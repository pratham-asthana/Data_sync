
import os
import json
from collections import defaultdict
from neo4j import GraphDatabase
from dotenv import load_dotenv

load_dotenv()

# -----------------------------
# Neo4j Configuration
# -----------------------------
NEO4J_URI = os.getenv("NEO4J_URI", "neo4j+s://2da6e3ab.databases.neo4j.io")
NEO4J_USERNAME = os.getenv("NEO4J_USERNAME", "neo4j")
NEO4J_PASSWORD = os.getenv("NEO4J_PASSWORD", "your-password")
NEO4J_DATABASE = os.getenv("NEO4J_DATABASE", "neo4j")

OUTPUT_FILE = "neo4j_schema.json"


class Neo4jSchemaExtractor:

    def __init__(self):
        self.driver = GraphDatabase.driver(
            NEO4J_URI,
            auth=(NEO4J_USERNAME, NEO4J_PASSWORD)
        )

    def close(self):
        self.driver.close()

    def get_schema(self):

        schema = {
            "statistics": {},
            "node_labels": {},
            "relationship_types": {},
            "relationship_mappings": []
        }

        with self.driver.session(database=NEO4J_DATABASE) as session:

            # -----------------------------
            # 1. Graph Statistics
            # -----------------------------
            stats = session.run("""
                MATCH (n)
                WITH count(n) AS node_count
                MATCH ()-[r]->()
                RETURN node_count, count(r) AS relationship_count
            """).single()

            schema["statistics"] = {
                "total_nodes": stats["node_count"],
                "total_relationships": stats["relationship_count"]
            }

            # -----------------------------
            # 2. Node Labels and Properties
            # -----------------------------
            labels_result = session.run("""
                CALL db.labels() YIELD label
                RETURN label
                ORDER BY label
            """)

            for record in labels_result:
                label = record["label"]

                properties_result = session.run("""
                    MATCH (n)
                    WHERE $label IN labels(n)
                    UNWIND keys(n) AS property
                    RETURN
                        property,
                        collect(DISTINCT valueType(n[property])) AS types,
                        count(*) AS occurrences
                    ORDER BY property
                """, label=label)

                properties = {}

                for prop in properties_result:
                    properties[prop["property"]] = {
                        "types": prop["types"],
                        "occurrences": prop["occurrences"]
                    }

                count_result = session.run("""
                    MATCH (n)
                    WHERE $label IN labels(n)
                    RETURN count(n) AS count
                """, label=label).single()

                schema["node_labels"][label] = {
                    "node_count": count_result["count"],
                    "properties": properties
                }

            # -----------------------------
            # 3. Relationship Types and Properties
            # -----------------------------
            relationship_result = session.run("""
                CALL db.relationshipTypes()
                YIELD relationshipType
                RETURN relationshipType
                ORDER BY relationshipType
            """)

            for record in relationship_result:
                rel_type = record["relationshipType"]

                properties_result = session.run("""
                    MATCH ()-[r]->()
                    WHERE type(r) = $rel_type
                    UNWIND keys(r) AS property
                    RETURN
                        property,
                        collect(DISTINCT valueType(r[property])) AS types,
                        count(*) AS occurrences
                    ORDER BY property
                """, rel_type=rel_type)

                properties = {}

                for prop in properties_result:
                    properties[prop["property"]] = {
                        "types": prop["types"],
                        "occurrences": prop["occurrences"]
                    }

                count_result = session.run("""
                    MATCH ()-[r]->()
                    WHERE type(r) = $rel_type
                    RETURN count(r) AS count
                """, rel_type=rel_type).single()

                schema["relationship_types"][rel_type] = {
                    "relationship_count": count_result["count"],
                    "properties": properties
                }

            # -----------------------------
            # 4. Relationship Mappings
            # -----------------------------
            mappings_result = session.run("""
                MATCH (source)-[r]->(target)
                UNWIND labels(source) AS source_label
                UNWIND labels(target) AS target_label
                RETURN DISTINCT
                    source_label,
                    type(r) AS relationship,
                    target_label
                ORDER BY source_label, relationship, target_label
            """)

            for record in mappings_result:
                schema["relationship_mappings"].append({
                    "source": record["source_label"],
                    "relationship": record["relationship"],
                    "target": record["target_label"]
                })

            # -----------------------------
            # 5. Additional Statistics
            # -----------------------------
            schema["statistics"].update({
                "total_labels": len(schema["node_labels"]),
                "total_relationship_types": len(
                    schema["relationship_types"]
                )
            })

        return schema


# -----------------------------
# Main Execution
# -----------------------------
if __name__ == "__main__":

    extractor = Neo4jSchemaExtractor()

    try:
        schema = extractor.get_schema()

        with open(OUTPUT_FILE, "w", encoding="utf-8") as file:
            json.dump(schema, file, indent=4, ensure_ascii=False)

        print("\nSchema extracted successfully!")
        print(f"Output: {OUTPUT_FILE}")

        print("\nGraph Statistics:")
        print(json.dumps(schema["statistics"], indent=4))

        print("\nNode Labels:")
        for label in schema["node_labels"]:
            print(f" - {label}")

        print("\nRelationship Types:")
        for rel in schema["relationship_types"]:
            print(f" - {rel}")

    except Exception as e:
        print(f"Error extracting schema: {e}")

    finally:
        extractor.close()
