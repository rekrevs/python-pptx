# Task T-{CATEGORY}-{NN}

## Header

| Field | Value |
|-------|-------|
| ID | T-{CATEGORY}-{NN} |
| Parent | B-{CATEGORY}-{NN} / T-{CATEGORY}-{NN} / User request |
| State | READY / IN_PROGRESS / BLOCKED / DONE / FAILED / PARTIAL |
| Created | YYYY-MM-DD |

## Objective

Brief description of what this task accomplishes.

## Acceptance Criteria

- [ ] Criterion 1
- [ ] Criterion 2
- [ ] Criterion 3

## Context

Reference to relevant:
- Existing code in `src/pptx/`
- OOXML analysis in `docs/dev/analysis/`
- Related tests in `tests/` or `features/`
- OOXML spec details

## Subtasks

| ID | Description | State |
|----|-------------|-------|
| T-{CATEGORY}-{NN}-1 | Subtask description | STATE |

## Implementation Notes

### Approach

Description of the implementation approach.

### Files Modified

- `path/to/file.py` - Description of changes

## Obstacles

### Obstacle 1: Description

- **Observed**: What happened
- **Expected**: What should happen
- **Tried**: What was attempted
- **Hypothesis**: Possible causes

## Evidence

### Tests Run

```
$ pytest tests/unit/test_xxx.py -v
... output ...
```

### Test Results

- Unit tests: PASS/FAIL
- Acceptance tests: PASS/FAIL

## Outcome

**State**: DONE / FAILED / BLOCKED / PARTIAL

Summary of what was accomplished or why task did not complete.
