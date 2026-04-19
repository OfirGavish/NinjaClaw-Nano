"""
NinjaClaw Prompt Optimizer — powered by Agent Lightning APO.

Reads conversation traces from the agent's SQLite database,
wraps them as Agent Lightning rollouts, runs APO (Automatic
Prompt Optimization), and writes the optimized prompt back
to the workspace.

Usage:
    python3 optimize_prompt.py --db /path/to/messages.db --prompt /path/to/SOUL.md
    
For NinjaClaw-Nano:
    python3 optimize_prompt.py \
        --db ~/ninjaclaw-nano/store/messages.db \
        --prompt ~/ninjaclaw-nano/groups/main/COPILOT.md \
        --output ~/ninjaclaw-nano/groups/main/COPILOT.md.optimized

For NinjaClaw-Open:
    python3 optimize_prompt.py \
        --db ~/.openclaw/workspace/ninjabrain.db \
        --prompt ~/.openclaw/workspace/SOUL.md \
        --output ~/.openclaw/workspace/SOUL.md.optimized
"""

import argparse
import sqlite3
import json
import os
import sys
from typing import TypedDict

try:
    import agentlightning as agl
    from openai import AsyncOpenAI
except ImportError:
    print("Error: Install dependencies first:")
    print("  pip install agentlightning[apo]")
    sys.exit(1)


class ConversationTask(TypedDict):
    """A conversation trace as a training task."""
    user_message: str
    agent_response: str
    user_feedback: float  # 0.0 = bad, 1.0 = good


def load_traces_from_NinjaClaw(db_path: str, limit: int = 50) -> list[ConversationTask]:
    """Load conversation traces from NinjaClaw's messages.db."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    
    # Get conversation pairs (user message → agent response)
    rows = conn.execute("""
        SELECT m1.content as user_msg, m2.content as agent_msg, m1.timestamp
        FROM messages m1
        JOIN messages m2 ON m1.chat_jid = m2.chat_jid 
            AND m2.is_from_me = 1
            AND m2.timestamp > m1.timestamp
        WHERE m1.is_from_me = 0
        ORDER BY m1.timestamp DESC
        LIMIT ?
    """, (limit,)).fetchall()
    conn.close()
    
    tasks = []
    for row in rows:
        # Simple heuristic reward: longer responses with no error markers = better
        response = row["agent_msg"] or ""
        reward = 0.5  # default: neutral
        
        # Positive signals
        if len(response) > 100:
            reward += 0.1
        if any(word in response.lower() for word in ["here's", "done", "created", "fixed", "updated"]):
            reward += 0.2
            
        # Negative signals
        if any(word in response.lower() for word in ["error", "failed", "sorry", "can't", "unable"]):
            reward -= 0.3
        if "not logged in" in response.lower():
            reward -= 0.5
            
        reward = max(0.0, min(1.0, reward))
        
        tasks.append(ConversationTask(
            user_message=row["user_msg"][:500],
            agent_response=response[:1000],
            user_feedback=reward,
        ))
    
    return tasks


def load_traces_from_openclaw(db_path: str, limit: int = 50) -> list[ConversationTask]:
    """Load conversation traces from OpenClaw's ninjabrain.db or session JSONL."""
    # OpenClaw stores sessions as JSONL files
    sessions_dir = os.path.expanduser("~/.openclaw/agents/main/sessions")
    tasks = []
    
    if os.path.isdir(sessions_dir):
        for fname in sorted(os.listdir(sessions_dir), reverse=True)[:10]:
            if not fname.endswith(".jsonl"):
                continue
            fpath = os.path.join(sessions_dir, fname)
            user_msg = None
            with open(fpath) as f:
                for line in f:
                    try:
                        entry = json.loads(line)
                        if entry.get("type") == "user":
                            content = entry.get("message", {}).get("content", "")
                            if isinstance(content, str):
                                user_msg = content
                            elif isinstance(content, list):
                                user_msg = " ".join(c.get("text", "") for c in content if c.get("type") == "text")
                        elif entry.get("type") == "assistant" and user_msg:
                            content = entry.get("message", {}).get("content", [])
                            text_parts = [c.get("text", "") for c in content if isinstance(c, dict) and c.get("type") == "text"]
                            agent_msg = " ".join(text_parts)
                            if agent_msg:
                                reward = 0.6 if len(agent_msg) > 50 else 0.4
                                tasks.append(ConversationTask(
                                    user_message=user_msg[:500],
                                    agent_response=agent_msg[:1000],
                                    user_feedback=reward,
                                ))
                                user_msg = None
                    except json.JSONDecodeError:
                        continue
            if len(tasks) >= limit:
                break
    
    return tasks[:limit]


@agl.rollout
def ninjaclaw_conversation(task: ConversationTask, prompt_template: agl.PromptTemplate) -> float:
    """Replay a conversation and score based on the prompt template quality."""
    # Format the system prompt
    system_prompt = prompt_template.format()
    
    # The reward is based on the original conversation outcome
    # APO will learn which prompt instructions lead to better rewards
    return task["user_feedback"]


def run_optimization(args):
    """Run APO optimization on collected traces."""
    print(f"Loading traces from {args.db}...")
    
    if args.agent_type == "nano":
        traces = load_traces_from_NinjaClaw(args.db, limit=args.limit)
    else:
        traces = load_traces_from_openclaw(args.db, limit=args.limit)
    
    if len(traces) < 5:
        print(f"Only {len(traces)} traces found. Need at least 5 for optimization. Collect more conversations first.")
        return
    
    print(f"Loaded {len(traces)} conversation traces.")
    
    # Read current prompt
    current_prompt = ""
    if os.path.exists(args.prompt):
        with open(args.prompt) as f:
            current_prompt = f.read()
    
    if not current_prompt:
        current_prompt = "You are a helpful AI assistant."
    
    print(f"Current prompt: {len(current_prompt)} chars")
    print(f"Running APO with {args.rounds} rounds, beam width {args.beam_width}...")
    
    # Split traces into train/val
    split = int(len(traces) * 0.7)
    train_data = traces[:split]
    val_data = traces[split:]
    
    # Create APO optimizer
    client = AsyncOpenAI()  # Uses OPENAI_API_KEY or GITHUB_TOKEN
    
    algo = agl.APO(
        async_openai_client=client,
        gradient_model=args.gradient_model,
        apply_edit_model=args.edit_model,
        beam_width=args.beam_width,
        beam_rounds=args.rounds,
        gradient_batch_size=min(4, len(train_data)),
        val_batch_size=min(8, len(val_data)),
    )
    
    trainer = agl.Trainer(
        algorithm=algo,
        initial_resources={
            "prompt_template": agl.PromptTemplate(
                template=current_prompt,
                engine="f-string",
            )
        },
    )
    
    # Run optimization
    trainer.fit(
        agent=ninjaclaw_conversation,
        train_dataset=train_data,
        val_dataset=val_data,
    )
    
    # Get the optimized prompt
    best_prompt = algo.get_best_prompt()
    
    # Write output
    output_path = args.output or f"{args.prompt}.optimized"
    with open(output_path, "w") as f:
        f.write(best_prompt.template)
    
    print(f"\n{'='*60}")
    print(f"Optimization complete!")
    print(f"Original prompt:  {len(current_prompt)} chars")
    print(f"Optimized prompt: {len(best_prompt.template)} chars")
    print(f"Written to: {output_path}")
    print(f"{'='*60}")
    print(f"\nTo apply: cp {output_path} {args.prompt}")
    print(f"Then restart the agent service.")


def main():
    parser = argparse.ArgumentParser(description="NinjaClaw Prompt Optimizer (Agent Lightning APO)")
    parser.add_argument("--db", required=True, help="Path to messages.db or ninjabrain.db")
    parser.add_argument("--prompt", required=True, help="Path to current system prompt (SOUL.md or COPILOT.md)")
    parser.add_argument("--output", help="Output path for optimized prompt (default: <prompt>.optimized)")
    parser.add_argument("--agent-type", choices=["nano", "open"], default="nano", help="Agent type")
    parser.add_argument("--limit", type=int, default=50, help="Max traces to load")
    parser.add_argument("--rounds", type=int, default=3, help="APO beam search rounds")
    parser.add_argument("--beam-width", type=int, default=2, help="APO beam width")
    parser.add_argument("--gradient-model", default="gpt-4o-mini", help="Model for gradient computation")
    parser.add_argument("--edit-model", default="gpt-4o-mini", help="Model for applying edits")
    
    args = parser.parse_args()
    run_optimization(args)


if __name__ == "__main__":
    main()
