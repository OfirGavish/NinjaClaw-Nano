/**
 * NinjaBrain — Structured knowledge engine for NinjaClaw.
 *
 * Compiled truth + timeline model. Uses NinjaClaw's existing better-sqlite3 db.
 * Entity types: person, company, concept, project, tool, other
 * Slug format: type/name (e.g. people/ofir-gavish, companies/microsoft)
 */

import Database from 'better-sqlite3';
import path from 'path';
import { STORE_DIR } from './config.js';
import { logger } from './logger.js';

const DB_PATH = path.join(STORE_DIR, 'ninjabrain.db');

let _db: Database.Database | null = null;

function getDb(): Database.Database {
  if (_db) return _db;
  _db = new Database(DB_PATH);
  _db.pragma('journal_mode = WAL');
  return _db;
}

export function initBrainSchema(): void {
  const db = getDb();
  db.exec(`
    CREATE TABLE IF NOT EXISTS brain_pages (
      slug TEXT PRIMARY KEY,
      type TEXT NOT NULL DEFAULT 'concept',
      title TEXT NOT NULL,
      compiled_truth TEXT NOT NULL DEFAULT '',
      timeline TEXT NOT NULL DEFAULT '',
      created_at REAL NOT NULL,
      updated_at REAL NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_brain_type ON brain_pages(type);

    CREATE TABLE IF NOT EXISTS brain_links (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      from_slug TEXT NOT NULL,
      to_slug TEXT NOT NULL,
      link_type TEXT NOT NULL DEFAULT 'references',
      created_at REAL NOT NULL,
      UNIQUE(from_slug, to_slug, link_type)
    );
    CREATE INDEX IF NOT EXISTS idx_brain_links_from ON brain_links(from_slug);
    CREATE INDEX IF NOT EXISTS idx_brain_links_to ON brain_links(to_slug);
  `);

  // FTS5 virtual table for full-text search
  try {
    db.exec(`
      CREATE VIRTUAL TABLE IF NOT EXISTS brain_fts USING fts5(
        slug, title, compiled_truth, timeline,
        content=brain_pages, content_rowid=rowid
      );
      CREATE TRIGGER IF NOT EXISTS brain_fts_insert AFTER INSERT ON brain_pages BEGIN
        INSERT INTO brain_fts(rowid, slug, title, compiled_truth, timeline)
        VALUES (new.rowid, new.slug, new.title, new.compiled_truth, new.timeline);
      END;
      CREATE TRIGGER IF NOT EXISTS brain_fts_update AFTER UPDATE ON brain_pages BEGIN
        INSERT INTO brain_fts(brain_fts, rowid, slug, title, compiled_truth, timeline)
        VALUES ('delete', old.rowid, old.slug, old.title, old.compiled_truth, old.timeline);
        INSERT INTO brain_fts(rowid, slug, title, compiled_truth, timeline)
        VALUES (new.rowid, new.slug, new.title, new.compiled_truth, new.timeline);
      END;
      CREATE TRIGGER IF NOT EXISTS brain_fts_delete AFTER DELETE ON brain_pages BEGIN
        INSERT INTO brain_fts(brain_fts, rowid, slug, title, compiled_truth, timeline)
        VALUES ('delete', old.rowid, old.slug, old.title, old.compiled_truth, old.timeline);
      END;
    `);
  } catch {
    // FTS already exists
  }
  logger.info('NinjaBrain schema initialized');
}

// --- Helpers ---

function normalizeSlug(slug: string): string {
  return slug.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9/_-]/g, '');
}

function typeFromSlug(slug: string): string {
  const map: Record<string, string> = {
    people: 'person', companies: 'company', concepts: 'concept',
    projects: 'project', tools: 'tool',
  };
  return map[slug.split('/')[0]] ?? 'other';
}

// --- Core operations ---

export function brainSearch(query: string, limit = 5, typeFilter = ''): Record<string, unknown>[] {
  const db = getDb();
  const sql = typeFilter
    ? `SELECT p.slug, p.type, p.title, p.compiled_truth, p.timeline, p.updated_at
       FROM brain_fts f JOIN brain_pages p ON f.slug = p.slug
       WHERE brain_fts MATCH ? AND p.type = ? ORDER BY rank LIMIT ?`
    : `SELECT p.slug, p.type, p.title, p.compiled_truth, p.timeline, p.updated_at
       FROM brain_fts f JOIN brain_pages p ON f.slug = p.slug
       WHERE brain_fts MATCH ? ORDER BY rank LIMIT ?`;
  const params = typeFilter ? [query, typeFilter, limit] : [query, limit];
  return db.prepare(sql).all(...params) as Record<string, unknown>[];
}

export function brainGet(slug: string): Record<string, unknown> | null {
  const db = getDb();
  const row = db.prepare('SELECT * FROM brain_pages WHERE slug = ?').get(slug) as Record<string, unknown> | undefined;
  if (!row) return null;
  const linksTo = db.prepare('SELECT to_slug, link_type FROM brain_links WHERE from_slug = ?').all(slug);
  const linksFrom = db.prepare('SELECT from_slug, link_type FROM brain_links WHERE to_slug = ?').all(slug);
  return { ...row, links_to: linksTo, links_from: linksFrom };
}

export function brainPut(slug: string, title: string, compiledTruth = '', timelineEntry = '', entityType = ''): { slug: string; action: string } {
  slug = normalizeSlug(slug);
  if (!entityType) entityType = typeFromSlug(slug);
  const now = Date.now() / 1000;
  const db = getDb();
  const existing = db.prepare('SELECT * FROM brain_pages WHERE slug = ?').get(slug) as Record<string, unknown> | undefined;

  if (existing) {
    const newTruth = compiledTruth || (existing.compiled_truth as string);
    let newTimeline = existing.timeline as string;
    if (timelineEntry) {
      const date = new Date().toISOString().slice(0, 10);
      newTimeline = `${newTimeline}\n- ${date}: ${timelineEntry}`.trim();
    }
    db.prepare('UPDATE brain_pages SET title=?, compiled_truth=?, timeline=?, updated_at=? WHERE slug=?')
      .run(title || (existing.title as string), newTruth, newTimeline, now, slug);
    return { slug, action: 'updated' };
  }

  const timeline = timelineEntry ? `- ${new Date().toISOString().slice(0, 10)}: ${timelineEntry}` : '';
  db.prepare('INSERT INTO brain_pages (slug, type, title, compiled_truth, timeline, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
    .run(slug, entityType, title, compiledTruth, timeline, now, now);
  return { slug, action: 'created' };
}

export function brainLink(fromSlug: string, toSlug: string, linkType = 'references'): string {
  fromSlug = normalizeSlug(fromSlug);
  toSlug = normalizeSlug(toSlug);
  const db = getDb();
  for (const s of [fromSlug, toSlug]) {
    if (!db.prepare('SELECT 1 FROM brain_pages WHERE slug = ?').get(s)) return `[ERROR] Page not found: ${s}`;
  }
  try {
    db.prepare('INSERT INTO brain_links (from_slug, to_slug, link_type, created_at) VALUES (?, ?, ?, ?)')
      .run(fromSlug, toSlug, linkType, Date.now() / 1000);
    return `Linked ${fromSlug} → ${toSlug} (${linkType})`;
  } catch { return `Link already exists: ${fromSlug} → ${toSlug} (${linkType})`; }
}

export function brainList(entityType = '', limit = 20): Record<string, unknown>[] {
  const db = getDb();
  const sql = entityType
    ? 'SELECT slug, type, title, updated_at FROM brain_pages WHERE type = ? ORDER BY updated_at DESC LIMIT ?'
    : 'SELECT slug, type, title, updated_at FROM brain_pages ORDER BY updated_at DESC LIMIT ?';
  return entityType ? db.prepare(sql).all(entityType, limit) as Record<string, unknown>[] : db.prepare(sql).all(limit) as Record<string, unknown>[];
}

export function brainDelete(slug: string): string {
  const db = getDb();
  if (!db.prepare('SELECT 1 FROM brain_pages WHERE slug = ?').get(slug)) return `[ERROR] Page not found: ${slug}`;
  db.prepare('DELETE FROM brain_links WHERE from_slug = ? OR to_slug = ?').run(slug, slug);
  db.prepare('DELETE FROM brain_pages WHERE slug = ?').run(slug);
  return `Deleted ${slug}`;
}

export function brainStats(): { total_pages: number; total_links: number; by_type: Record<string, number> } {
  const db = getDb();
  const total = (db.prepare('SELECT COUNT(*) as c FROM brain_pages').get() as { c: number }).c;
  const links = (db.prepare('SELECT COUNT(*) as c FROM brain_links').get() as { c: number }).c;
  const byType = db.prepare('SELECT type, COUNT(*) as c FROM brain_pages GROUP BY type ORDER BY c DESC').all() as { type: string; c: number }[];
  return { total_pages: total, total_links: links, by_type: Object.fromEntries(byType.map(r => [r.type, r.c])) };
}

/** Search brain for entities mentioned in user text, return context string. */
export function brainContextForMessage(userText: string): string {
  if (!userText || userText.length < 3) return '';
  try {
    const results = brainSearch(userText.slice(0, 100), 3);
    if (results.length === 0) return '';
    const parts = ['## NinjaBrain Context\n'];
    for (const p of results) {
      parts.push(`### ${p.title} (${p.type}: ${p.slug})`);
      if (p.compiled_truth) parts.push(p.compiled_truth as string);
      if (p.timeline) parts.push(`Timeline:\n${p.timeline}`);
      parts.push('');
    }
    return parts.join('\n');
  } catch { return ''; }
}
