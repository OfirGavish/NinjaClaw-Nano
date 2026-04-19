#!/usr/bin/env bash
set -euo pipefail

# ============================================================
# setup-agent-lightning.sh — Install Agent Lightning on a NinjaClaw VM
#
# Usage (run on the VM as root or with sudo):
#   bash setup-agent-lightning.sh nano   # for NinjaClaw-Nano
#   bash setup-agent-lightning.sh open   # for NinjaClaw-Open
# ============================================================

AGENT_TYPE="${1:-nano}"
USER="azureuser"
HOME_DIR="/home/$USER"
VENV_DIR="$HOME_DIR/agl-env"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Setting up Agent Lightning for NinjaClaw-${AGENT_TYPE} ==="

# Step 1: Install system dependencies
echo ">>> Installing python3-venv..."
apt-get install -qq -y python3.10-venv python3-pip sqlite3 2>/dev/null || \
apt-get install -qq -y python3-venv python3-pip sqlite3

# Step 2: Create Python virtual environment
if [ ! -d "$VENV_DIR" ]; then
    echo ">>> Creating Python venv at $VENV_DIR..."
    su - "$USER" -c "python3 -m venv $VENV_DIR"
fi

# Step 3: Install Agent Lightning
echo ">>> Installing Agent Lightning..."
su - "$USER" -c "source $VENV_DIR/bin/activate && pip install -q agentlightning"

# Verify
AGL_VERSION=$(su - "$USER" -c "source $VENV_DIR/bin/activate && python3 -c 'import agentlightning; print(agentlightning.__version__)' 2>&1")
echo ">>> Agent Lightning version: $AGL_VERSION"

# Step 4: Deploy optimizer script
echo ">>> Deploying optimize_prompt.py..."
cp "$SCRIPT_DIR/optimize_prompt.py" "$HOME_DIR/optimize_prompt.py"
chown "$USER:$USER" "$HOME_DIR/optimize_prompt.py"
chmod +x "$HOME_DIR/optimize_prompt.py"

# Step 5: Set up cron job (daily at 3 AM)
echo ">>> Setting up daily optimization cron job..."

if [ "$AGENT_TYPE" = "nano" ]; then
    DB_PATH="$HOME_DIR/ninjaclaw-nano/store/messages.db"
    PROMPT_PATH="$HOME_DIR/ninjaclaw-nano/groups/main/COPILOT.md"
elif [ "$AGENT_TYPE" = "open" ]; then
    DB_PATH="$HOME_DIR/.openclaw/workspace/ninjabrain.db"
    PROMPT_PATH="$HOME_DIR/.openclaw/workspace/SOUL.md"
else
    echo "ERROR: Unknown agent type '$AGENT_TYPE'. Use 'nano' or 'open'."
    exit 1
fi

CRON_CMD="0 3 * * * $VENV_DIR/bin/python3 $HOME_DIR/optimize_prompt.py --db $DB_PATH --prompt $PROMPT_PATH --agent-type $AGENT_TYPE >> $HOME_DIR/agl-optimize.log 2>&1"

su - "$USER" -c "echo '$CRON_CMD' | crontab -"
echo ">>> Cron job installed:"
su - "$USER" -c "crontab -l"

echo ""
echo "==========================================================="
echo "  Agent Lightning Setup Complete"
echo "==========================================================="
echo ""
echo "  Agent type:    $AGENT_TYPE"
echo "  Venv:          $VENV_DIR"
echo "  Optimizer:     $HOME_DIR/optimize_prompt.py"
echo "  DB:            $DB_PATH"
echo "  Prompt:        $PROMPT_PATH"
echo "  Cron:          Daily at 3 AM"
echo "  Log:           $HOME_DIR/agl-optimize.log"
echo ""
echo "  Manual run:"
echo "    source $VENV_DIR/bin/activate"
echo "    python3 $HOME_DIR/optimize_prompt.py \\"
echo "      --db $DB_PATH \\"
echo "      --prompt $PROMPT_PATH \\"
echo "      --agent-type $AGENT_TYPE"
echo ""
echo "  After optimization, review and apply:"
echo "    diff $PROMPT_PATH ${PROMPT_PATH}.optimized"
echo "    cp ${PROMPT_PATH}.optimized $PROMPT_PATH"
echo "    # Then restart the agent service"
echo ""
echo "==========================================================="
