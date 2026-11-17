#!/usr/bin/env python3
"""
Neo4j Knowledge Graph Document Ingestion Script

This script ingests document metadata JSON into the Neo4j knowledge graph with:
- Upsert behavior: merges on sharepoint_url (unique key), replaces all properties if exists
- Dual labels: Document + specific type (RTM, SDD, RDD, USER_GUIDE, etc.)
- Proper hierarchy: Company -> Track -> Project -> Document
- Relationships: HAS_TRACK (Company->Track), HAS_DOCUMENT (Track->Document, Project->Document)
- No duplicates: sharepoint_url is the unique constraint key

Usage:
    python ingest_documents.py path/to/documents.json
    python ingest_documents.py path/to/documents.json --dry-run
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any


# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURATION - Matches existing graph schema
# ──────────────────────────────────────────────────────────────────────────────

COMPANY_NAME = "Orbrick"

# Valid document types (must match existing node labels in the graph)
VALID_DOC_TYPES = {
    "RTM", "SDD", "RDD", "USER_GUIDE", "DATA_MIGRATION", "PROCESS_FLOW",
    "CONFIG_WORKBOOK", "TEST_SCRIPT", "HLA", "TDD", "DOA", "AUDIT",
    "ORACLE_SR", "REPORT_FORMAT", "REPORT_INVENTORY", "REPORT_REVIEW",
    "REPORT_TEST_OUTPUT", "KICKOFF_DECK", "TEMPLATE", "QUESTIONNAIRE",
    "PROCESS_DISCOVERY", "POC", "MIGRATION_DOC", "TICKET_RCA", "TICKET_LOG",
    "SAMPLE_DATA", "TRAINING", "STATUS_REPORT", "SOW", "MOM", "UNKNOWN",
    "FAST_FORMULA", "RAID_LOG", "HANDOVER",
}

# Track names that exist in the graph
EXISTING_TRACKS = {"HCM", "SCM"}

# Project names that exist in the graph
EXISTING_PROJECTS = {
    "Blue Ocean Health", "Emicool", "Hendrick Health", "Network Intelligence",
    "Plus Malaysia", "TRPGlobal System Audit-AD Ports", "Tawuniya-PWC-Augmentation"
}

# Department to Track mapping
DEPARTMENT_TO_TRACK = {
    "HCM": "HCM",
    "SCM": "SCM",
    "FIN": "FIN",  # may need to create
    "EPM": "EPM",  # may need to create
    "Extensions": "Extensions",  # may need to create
}

# Project inference from document path/name (for connecting to Project nodes)
# This maps path keywords to project names
PROJECT_KEYWORDS = {
    "blue ocean": "Blue Ocean Health",
    "boh": "Blue Ocean Health",
    "emicool": "Emicool",
    "hendrick": "Hendrick Health",
    "network intelligence": "Network Intelligence",
    "ni": "Network Intelligence",
    "plus malaysia": "Plus Malaysia",
    "trpglobal": "TRPGlobal System Audit-AD Ports",
    "trp": "TRPGlobal System Audit-AD Ports",
    "tawuniya": "Tawuniya-PWC-Augmentation",
    "pwc": "Tawuniya-PWC-Augmentation",
    "yashraj": "Blue Ocean Health",  # YASHRAJ appears to be under Blue Ocean Health
}


# ──────────────────────────────────────────────────────────────────────────────
# CYPHER QUERY TEMPLATES
# ──────────────────────────────────────────────────────────────────────────────

def build_ensure_hierarchy_cypher() -> str:
    """Create Company, Tracks, Projects and their relationships if not exist."""
    track_merges = []
    for track in sorted(EXISTING_TRACKS):
        track_merges.append(f"MERGE (t_{track.lower()}:Track {{name: '{track}'}})")
        track_merges.append(f"MERGE (c)-[:HAS_TRACK]->(t_{track.lower()})")
    
    project_merges = []
    for project in sorted(EXISTING_PROJECTS):
        safe_name = project.lower().replace(" ", "_").replace("-", "_").replace(".", "")
        project_merges.append(f"MERGE (p_{safe_name}:Project {{name: '{project}'}})")
        project_merges.append(f"MERGE (c)-[:HAS_TRACK]->(p_{safe_name})")
        # Projects also connect to tracks - we'll link based on department later
    
    return f"""
// Ensure Company exists
MERGE (c:Company {{name: '{COMPANY_NAME}'}})

// Ensure Tracks exist and connect to Company
{chr(10).join(track_merges)}

// Ensure Projects exist and connect to Company
{chr(10).join(project_merges)}
"""


def infer_project_from_path(sharepoint_path: str, document_name: str) -> str | None:
    """Infer project name from SharePoint path or document name."""
    text = f"{sharepoint_path} {document_name}".lower()
    for keyword, project in PROJECT_KEYWORDS.items():
        if keyword in text:
            return project
    return None


def build_document_upsert_cypher(doc: dict[str, Any], track_var: str, project_var: str | None) -> str:
    """
    Build Cypher to upsert a document node.
    Uses sharepoint_url as unique key (MERGE on sharepoint_url).
    Dual labels: Document + specific type (e.g., Document:RTM)
    """
    doc_type = doc.get("document_type", "UNKNOWN")
    if doc_type not in VALID_DOC_TYPES:
        doc_type = "UNKNOWN"
    
    # Escape single quotes for Cypher string literals
    def esc(v: str) -> str:
        if v is None:
            return ""
        return str(v).replace("'", "\\'").replace("\n", "\\n").replace("\r", "\\r")
    
    props = {
        "department": esc(doc.get("department", "")),
        "document_name": esc(doc.get("document_name", "")),
        "document_type": esc(doc.get("document_type", "")),
        "file_extension": esc(doc.get("file_extension", "")),
        "modified_timestamp": esc(doc.get("modified_timestamp", "")),
        "sharepoint_path": esc(doc.get("sharepoint_path", "")),
        "sharepoint_url": esc(doc.get("sharepoint_url", "")),
        "version": esc(doc.get("version", "UNKNOWN")),
        "summary": esc(doc.get("summary", "")),
    }
    
    # Build SET clause - update ALL properties on match
    set_clauses = ",\n    ".join(f"d.{k} = '{v}'" for k, v in props.items())
    
    # Build relationship merges
    rel_clauses = [f"MERGE ({track_var})-[:HAS_DOCUMENT]->(d)"]
    if project_var:
        rel_clauses.append(f"MERGE ({project_var})-[:HAS_DOCUMENT]->(d)")
    
    return f"""
// Upsert document: {props['document_name']}
MERGE (d:Document:`{doc_type}` {{sharepoint_url: '{props['sharepoint_url']}'}})
SET
    {set_clauses}
{chr(10).join(rel_clauses)}
"""


def build_cleanup_cypher() -> str:
    """Remove any duplicate HAS_DOCUMENT relationships."""
    return """
// Clean up duplicate HAS_DOCUMENT relationships (keep one per track/document pair)
MATCH (t:Track)-[r:HAS_DOCUMENT]->(d:Document)
WITH t, d, collect(r) as rels
WHERE size(rels) > 1
UNWIND rels[1..] as extra_rel
DELETE extra_rel

// Clean up duplicate HAS_DOCUMENT relationships (keep one per project/document pair)
MATCH (p:Project)-[r:HAS_DOCUMENT]->(d:Document)
WITH p, d, collect(r) as rels
WHERE size(rels) > 1
UNWIND rels[1..] as extra_rel
DELETE extra_rel
"""


def build_verification_cypher() -> str:
    """Verification queries to run after ingestion."""
    return """
// Verification: Document counts by track and type
MATCH (t:Track)-[:HAS_DOCUMENT]->(d:Document)
RETURN t.name as track, labels(d)[1] as doc_type, count(*) as count
ORDER BY track, count DESC

// Verification: Document counts by project and type
MATCH (p:Project)-[:HAS_DOCUMENT]->(d:Document)
RETURN p.name as project, labels(d)[1] as doc_type, count(*) as count
ORDER BY project, count DESC

// Verification: Total documents
MATCH (d:Document) RETURN count(*) as total_documents

// Verification: Tracks and Projects
MATCH (c:Company)-[:HAS_TRACK]->(t:Track) RETURN c.name as company, collect(t.name) as tracks
MATCH (c:Company)-[:HAS_TRACK]->(p:Project) RETURN c.name as company, collect(p.name) as projects
"""


# ──────────────────────────────────────────────────────────────────────────────
# MAIN INGESTION LOGIC
# ──────────────────────────────────────────────────────────────────────────────

def load_json_file(filepath: Path) -> list[dict]:
    """Load and validate JSON file."""
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    if not isinstance(data, list):
        raise ValueError(f"Expected JSON array, got {type(data).__name__}")
    
    valid_records = []
    for i, record in enumerate(data):
        if not isinstance(record, dict):
            print(f"  ⚠ Skipping record {i}: not an object")
            continue
        
        # Required fields check
        required = ["sharepoint_url", "document_name", "document_type", "department"]
        missing = [f for f in required if f not in record or not record[f]]
        if missing:
            print(f"  ⚠ Skipping record {i}: missing required fields {missing}")
            continue
        
        # Normalize version
        if not record.get("version") or record["version"].strip() == "":
            record["version"] = "UNKNOWN"
        
        # Normalize document_type
        if record["document_type"] not in VALID_DOC_TYPES:
            print(f"  ⚠ Record {i}: unknown document_type '{record['document_type']}', using UNKNOWN")
            record["document_type"] = "UNKNOWN"
        
        valid_records.append(record)
    
    return valid_records


def generate_cypher_script(documents: list[dict], output_path: Path | None = None) -> str:
    """Generate complete Cypher script for ingestion."""
    lines = [
        "// ═══════════════════════════════════════════════════════════════",
        "// Neo4j Document Ingestion Script",
        f"// Generated for {len(documents)} documents",
        "// ═══════════════════════════════════════════════════════════════",
        "",
        "// 1. Ensure Company, Tracks, Projects hierarchy exists",
        build_ensure_hierarchy_cypher().strip(),
        "",
        "// 2. Upsert documents (deduplication via sharepoint_url)",
    ]
    
    for doc in documents:
        department = doc.get("department", "HCM")
        track_var = f"t_{department.lower()}"
        
        # Infer project from path
        project_name = infer_project_from_path(
            doc.get("sharepoint_path", ""), 
            doc.get("document_name", "")
        )
        project_var = None
        if project_name:
            safe_name = project_name.lower().replace(" ", "_").replace("-", "_").replace(".", "")
            project_var = f"p_{safe_name}"
        
        lines.append(build_document_upsert_cypher(doc, track_var, project_var).strip())
    
    lines.extend([
        "",
        "// 3. Clean up any duplicate relationships",
        build_cleanup_cypher().strip(),
        "",
        "// 4. Verification reports",
        build_verification_cypher().strip(),
    ])
    
    cypher = "\n\n".join(lines)
    
    if output_path:
        output_path.write_text(cypher, encoding="utf-8")
        print(f"  ✓ Cypher script written to: {output_path}")
    
    return cypher


def execute_via_mcp(cypher: str) -> bool:
    """Execute Cypher via the Neo4j MCP tool (placeholder - would need actual MCP integration)."""
    # This would be called via the mcp__neo4j__write_cypher tool
    # For now, we just return the script for manual execution
    return True


def print_execution_instructions(cypher_path: Path):
    """Print instructions for executing the Cypher."""
    print("\n" + "=" * 70)
    print("EXECUTION INSTRUCTIONS")
    print("=" * 70)
    print(f"""
Option 1: Via Hermes MCP (recommended)
  The script can be executed using the mcp__neo4j__write_cypher tool.
  In Hermes, use the tool with the query parameter containing the Cypher.

Option 2: Via Neo4j Browser
  1. Open Neo4j Browser (http://localhost:7474)
  2. Copy contents of {cypher_path.name}
  3. Paste and click Run

Option 3: Via Cypher Shell
  cat {cypher_path} | cypher-shell -u neo4j -p <password>

Option 4: Via Python Neo4j Driver
  from neo4j import GraphDatabase
  driver = GraphDatabase.driver("bolt://localhost:7687", auth=("neo4j", "password"))
  with driver.session() as session:
      with open("{cypher_path}") as f:
          # Split by ';' and execute each statement
          for stmt in f.read().split(';'):
              if stmt.strip():
                  session.run(stmt)

IMPORTANT: The script uses MERGE on sharepoint_url as unique key.
- If document exists (same sharepoint_url): ALL properties are UPDATED with new values
- If document doesn't exist: NEW node created with dual label (Document + type)
- Relationships to Track and Project are created/updated accordingly
""")


def main():
    parser = argparse.ArgumentParser(
        description="Generate Cypher script to ingest JSON documents into Neo4j knowledge graph",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python ingest_documents.py "YASHRAJ 1.json"
  python ingest_documents.py "YASHRAJ 1.json" --output ingest.cypher
  python ingest_documents.py ./new_site_exports/ --recursive
  python ingest_documents.py "YASHRAJ 1.json" --dry-run
        """
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="JSON file(s) or directory(ies) to ingest"
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=Path("ingest_documents.cypher"),
        help="Output Cypher script path (default: ingest_documents.cypher)"
    )
    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="Recursively search directories for .json files"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only validate and report, don't generate Cypher"
    )
    
    args = parser.parse_args()
    
    # Collect all JSON files
    json_files = []
    for input_path in args.inputs:
        if input_path.is_dir():
            pattern = "**/*.json" if args.recursive else "*.json"
            json_files.extend(input_path.glob(pattern))
        elif input_path.suffix == ".json":
            json_files.append(input_path)
        else:
            print(f"⚠ Skipping {input_path}: not a .json file or directory")
    
    if not json_files:
        print("❌ No JSON files found")
        sys.exit(1)
    
    print(f"📂 Found {len(json_files)} JSON file(s)")
    
    # Load all documents
    all_documents = []
    for jf in json_files:
        print(f"  Loading {jf.name}...")
        docs = load_json_file(jf)
        print(f"    ✓ {len(docs)} valid document(s)")
        all_documents.extend(docs)
    
    print(f"\n📄 Total documents to ingest: {len(all_documents)}")
    
    # Summary by type/department
    from collections import Counter
    type_counts = Counter(d["document_type"] for d in all_documents)
    dept_counts = Counter(d["department"] for d in all_documents)
    
    print("\n📊 By document_type:")
    for t, c in sorted(type_counts.items()):
        print(f"    {t}: {c}")
    
    print("\n📊 By department:")
    for d, c in sorted(dept_counts.items()):
        print(f"    {d}: {c}")
    
    # Check for duplicates in input (same sharepoint_url)
    url_counts = Counter(d["sharepoint_url"] for d in all_documents)
    dupes = {url: cnt for url, cnt in url_counts.items() if cnt > 1}
    if dupes:
        print(f"\n⚠️  WARNING: {len(dupes)} duplicate sharepoint_url(s) in input:")
        for url, cnt in dupes.items():
            print(f"    {cnt}x: {url[:80]}...")
        print("    (Last occurrence will win due to MERGE behavior)")
    
    if args.dry_run:
        print("\n✅ Dry run complete - no Cypher generated")
        return
    
    # Generate Cypher
    print(f"\n🔧 Generating Cypher script...")
    cypher = generate_cypher_script(all_documents, args.output)
    
    print(f"\n✅ Cypher script generated: {args.output}")
    print(f"   Approximate statements: {cypher.count('MERGE') + cypher.count('MATCH')} operations")
    
    print_execution_instructions(args.output)


if __name__ == "__main__":
    main()