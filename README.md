## Architecture

Raw Daily Files (8 Partners)
→ Power Query Standardization
→ Monthly Partner Dataset
→ Consolidated Master Dataset
→ VBA Orchestration with Logging

## Logging System

Each refresh run generates:
- Run ID
- Start & End Time
- Duration (seconds)
- File-level status (SUCCESS/FAILED)
- Error message (if any)
