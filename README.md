# SharePoint → Neo4j Data Sync Pipeline

## Overview

This pipeline synchronizes SharePoint document metadata to a Neo4j knowledge graph, maintaining a structured hierarchy of Company → Track → Project → Document with full deduplication and upsert capabilities.

---

## Architecture Diagram

```mermaid
flowchart TD
    %% Styles
    classDef source fill:#e3f2fd,stroke:#1976d2,stroke-width:2px
    classDef process fill:#fff3e0,stroke:#f57c00,stroke-width:2px
    classDef storage fill:#e8f5e9,stroke:#388e3c,stroke-width:2px
    classDef output fill:#fce4ec,stroke:#c2185b,stroke-width:2px
    classDef decision fill:#fff9c4,stroke:#fbc02d,stroke-width:2px,dotted

    %% Source Layer
    SP[("SharePoint Sites\norbrick.sharepoint.com")]:::source
    ENV[(".env credentials")]:::source

    %% Comparison Layer
    LIST[list_sharepoint_sites.py]:::process
    COMP[comparison.py]:::process
    CACHE[(sharepoint_sites.json\nprevious snapshot)]:::storage

    %% Pipeline Layer
    PIPE[sharepoint_new_site_pipeline.py]:::process
    GRAPH[GraphClient\n(sharepoint_items_json_with_date.py)]:::process
    
    %% Document Classification
    CLASS[Document Classification Engine]:::process
    TRACKS[TRACKS mapping\n(hcm→HCM, scm→SCM, etc.)]:::storage
    DOCPATHS[DOCUMENT_PATHS patterns\n(understudy/rtm→RTM, etc.)]:::storage

    %% Output Generation
    EXPORT[Export Site Data]:::process
    STRUCT[structure.txt\n(folder tree)]:::output
    JSON[<site_name>.json\n(document records)]:::output

    %% Neo4j Ingestion
    MAIN[main.py\n(Orchestrator)]:::process
    INGEST[ingest_to_neo4j.py]:::process
    CYPHER[ingest_generated.cypher]:::output
    
    %% Neo4j Database
    NEO4J[("Neo4j Knowledge Graph")]:::storage
    
    %% Graph Schema
    SCHEMA[Neo4j Schema\n37 Document Types\n2 Tracks\n7 Projects]:::storage

    %% Flow
    SP --> LIST
    ENV --> LIST
    LIST --> COMP
    CACHE --> COMP
    COMP -->|added_sites| PIPE
    PIPE --> GRAPH
    GRAPH --> CLASS
    TRACKS --> CLASS
    DOCPATHS --> CLASS
    CLASS --> EXPORT
    EXPORT --> STRUCT
    EXPORT --> JSON
    JSON --> MAIN
    MAIN --> INGEST
    INGEST --> CYPHER
    CYPHER --> NEO4J
    SCHEMA --> NEO4J

    %% Subgraph for Neo4j Schema
    subgraph NEO4J_SCHEMA [Neo4j Graph Structure]
        COMPANY[Company: Orbrick]
        TRACK_HCM[Track: HCM]
        TRACK_SCM[Track: SCM]
        PROJ_BOH[Project: Blue Ocean Health]
        PROJ_EMI[Project: Emicool]
        PROJ_HEN[Project: Hendrick Health]
        PROJ_NI[Project: Network Intelligence]
        PROJ_PM[Project: Plus Malaysia]
        PROJ_TRP[Project: TRPGlobal]
        PROJ_TAW[Project: Tawuniya-PWC]
        DOC_RTM[Document:RTM]
        DOC_SDD[Document:SDD]
        DOC_RDD[Document:RDD]
        DOC_UG[Document:USER_GUIDE]
        DOC_DM[Document:DATA_MIGRATION]
        DOC_OTHER[...32 more types]
    end

    COMPANY --> TRACK_HCM
    COMPANY --> TRACK_SCM
    COMPANY -.-> PROJ_BOH
    COMPANY -.-> PROJ_EMI
    COMPANY -.-> PROJ_HEN
    COMPANY -.-> PROJ_NI
    COMPANY -.-> PROJ_PM
    COMPANY -.-> PROJ_TRP
    COMPANY -.-> PROJ_TAW
    TRACK_HCM -.-> DOC_RTM
    TRACK_HCM -.-> DOC_SDD
    TRACK_HCM -.-> DOC_RDD
    TRACK_HCM -.-> DOC_UG
    TRACK_HCM -.-> DOC_DM
    TRACK_HCM -.-> DOC_OTHER
    TRACK_SCM -.-> DOC_RTM
    TRACK_SCM -.-> DOC_SDD
    PROJ_BOH -.-> DOC_RTM
    PROJ_BOH -.-> DOC_SDD
```

---

## Component Details

### 1. SharePoint Discovery (`comparison.py`)

| Aspect | Detail |
|--------|--------|
| **Input** | Current SharePoint site list via `list_sharepoint_sites.py` |
| **State** | `sharepoint_sites.json` (previous snapshot) |
| **Output** | `added_sites` array with `{name, url, id, displayName}` |
| **Key Logic** | Compares site IDs and URLs to detect new sites |

### 2. Document Extraction (`sharepoint_new_site_pipeline.py`)

| Aspect | Detail |
|--------|--------|
| **Authentication** | `.env` → `get_access_token()` → `GraphClient` |
| **Traversal** | Recursive folder scan via `list_children()` / `list_drives()` |
| **Classification** | Path pattern matching → `department` + `document_type` |
| **Output** | Per-site JSON + `structure.txt` folder tree |

#### Classification Rules
```python
TRACKS = {
    "hcm": "HCM", "scm": "SCM", "fin": "FIN", 
    "epm": "EPM", "extensions": "Extensions"
}

DOCUMENT_PATHS = [
    (("understudy", "rtm"), "RTM"),
    (("redesign", "config workbooks"), "CONFIG_WORKBOOK"),
    (("redesign", "sdd"), "SDD"),
    (("orchestrate", "kut"), "USER_GUIDE"),
    (("orchestrate", "uat"), "USER_GUIDE"),
]
```

### 3. Document Record Schema

```json
{
  "department": "HCM",
  "document_name": "BOH_HCM_Absence_SDD_v2.0.docx",
  "document_type": "SDD",
  "file_extension": "docx",
  "modified_timestamp": "2026-11-02T07:38:27Z",
  "sharepoint_path": "/YASHRAJ/3. Redesign/SDD/HCM/...",
  "sharepoint_url": "https://orbrick.sharepoint.com/...",
  "version": "v2.0"
}
```

### 4. Neo4j Ingestion (`ingest_to_neo4j.py` / `main.py`)

| Feature | Implementation |
|---------|---------------|
| **Deduplication** | `MERGE (d:Document:\`TYPE\` {sharepoint_url})` |
| **Upsert** | `SET` all properties on match (replace with new) |
| **Dual Labels** | `Document:RTM`, `Document:SDD`, etc. |
| **Hierarchy** | Company → Track → Project → Document |
| **Relationships** | `HAS_TRACK`, `HAS_DOCUMENT` |
| **Project Inference** | Keyword matching on path/name |

#### Generated Cypher Structure
```cypher
// 1. Ensure hierarchy
MERGE (c:Company {name: 'Orbrick'})
MERGE (t_hcm:Track {name: 'HCM'}) MERGE (c)-[:HAS_TRACK]->(t_hcm)
MERGE (p_boh:Project {name: 'Blue Ocean Health'}) MERGE (c)-[:HAS_TRACK]->(p_boh)

// 2. Upsert documents
MERGE (d:Document:`RTM` {sharepoint_url: '...'})
SET d.department = 'HCM', d.document_name = '...', ...
MERGE (t_hcm)-[:HAS_DOCUMENT]->(d)
MERGE (p_boh)-[:HAS_DOCUMENT]->(d)

// 3. Cleanup duplicates
// 4. Verification queries
```

---

## Neo4j Graph Schema

### Node Labels (37 Document Types)

| Category | Types |
|----------|-------|
| **Requirements & Design** | RDD, SDD, TDD, HLA |
| **Testing** | TEST_SCRIPT, REPORT_TEST_OUTPUT, REPORT_REVIEW, SAMPLE_DATA |
| **Process & Config** | PROCESS_FLOW, CONFIG_WORKBOOK, PROCESS_DISCOVERY, FAST_FORMULA |
| **Migration & Data** | DATA_MIGRATION, MIGRATION_DOC |
| **User Enablement** | USER_GUIDE, TRAINING |
| **Reports** | REPORT_FORMAT, REPORT_INVENTORY, STATUS_REPORT |
| **Project Mgmt** | KICKOFF_DECK, RAID_LOG, SOW, MOM, HANDOVER |
| **Tickets & Issues** | TICKET_LOG, TICKET_RCA |
| **Other** | POC, DOA, AUDIT, ORACLE_SR, RTM, TEMPLATE, QUESTIONNAIRE, UNKNOWN, RAID_LOG, HANDOVER |

### Current Graph Stats
- **Company**: 1 (Orbrick)
- **Tracks**: 2 (HCM, SCM)
- **Projects**: 7
- **Documents**: ~391 (growing)
- **Relationships**: ~755 HAS_DOCUMENT, 9 HAS_TRACK

---

## Usage

### Full Pipeline Sync
```bash
cd /c/Users/Lenovo/Desktop/sync
python main.py
```

### Dry Run (Generate Cypher Only)
```bash
python main.py --dry-run
```

### Ingest Existing JSON Files
```bash
python main.py --skip-sharepoint --output-dir "C:/path/to/json"
```

### Execute Cypher via Hermes MCP
```bash
python main.py --execute-cypher
```

---

## File Structure

```
Desktop/sync/
├── main.py                          # Orchestrator (ENTRY POINT)
├── comparison.py                    # Site diff detection
├── sharepoint_new_site_pipeline.py  # Document extraction & classification
├── ingest_to_neo4j.py               # Cypher generation for Neo4j
├── list_sharepoint_sites.py         # SharePoint API: list sites
├── sharepoint_items_json_with_date.py # SharePoint API: traverse folders
├── sharepoint_sites.json            # Persistent site registry
├── new_site_exports/                # Pipeline output directory
│   └── <site_name>/
│       ├── <site_name>.json         # Document records
│       └── structure.txt            # Folder tree visualization
├── ingest_generated.cypher          # Generated Neo4j Cypher
├── neo4j_schema.json                # Exported graph schema
└── .env                             # SharePoint credentials
```

---

## Key Design Decisions

1. **sharepoint_url as Primary Key** - Globally unique, enables safe upserts
2. **Dual Labels** - `Document:TYPE` allows both polymorphic and type-specific queries
3. **Track-Level Classification** - Department maps to Track, not Project
4. **Project Inference** - Heuristic keyword matching (extensible via `PROJECT_KEYWORDS`)
5. **Immutable History** - `comparison.py` maintains site registry for incremental sync
6. **No Summary Field** - Excluded per requirement (separate processing pipeline)

---

## Extending the Pipeline

### Add New Document Type
1. Add to `VALID_DOC_TYPES` in `ingest_to_neo4j.py`
2. Add path pattern to `DOCUMENT_PATHS` in `sharepoint_new_site_pipeline.py`
3. Ensure Neo4j has the label (created automatically via `MERGE`)

### Add New Track
1. Add to `TRACKS` mapping in pipeline
2. Add to `EXISTING_TRACKS` in ingestion
3. Add to `DEPARTMENT_TO_TRACK` mapping

### Add New Project
1. Add to `EXISTING_PROJECTS` in ingestion
2. Add keyword to `PROJECT_KEYWORDS` for auto-inference

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| No new sites detected | Check `sharepoint_sites.json` isn't stale; delete to force full refresh |
| Classification fails | Verify folder names match `DOCUMENT_PATHS` patterns (case-insensitive) |
| Cypher timeout | Use `--dry-run` to inspect; split large batches manually |
| Duplicate documents | `sharepoint_url` MERGE handles this; check for URL variations |
| Missing project links | Add keywords to `PROJECT_KEYWORDS` mapping |