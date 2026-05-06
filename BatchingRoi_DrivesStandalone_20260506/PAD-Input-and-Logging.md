# Input and Logging Design

## Accepted Input Lengths
- `8`: enter only, no tab.
- `4` or `7`: enter, then tab.
- Other lengths: log exception, skip.

## Input Source Options

### Option A: Manual list variable
Create `BatchList` directly in PAD as a list of text values.

### Option B: Text file
1. `Read text from file` (one number per line).
2. `Split text` by new line -> `BatchList`.
3. Reuse the same loop logic.

### Option C: Excel
1. `Launch Excel` and `Read from worksheet`.
2. Convert first column to list.
3. Reuse the same loop logic.

## Exception Log CSV Format
Header:

```text
Timestamp,Value,Reason
```

Example rows:

```text
2026-05-06 10:12:03,99,Invalid length (2)
2026-05-06 10:12:04,123456789,Invalid length (9)
```

## Optional Extra Validation
Before length check, strip spaces and validate digits-only using regex.

Suggested rule:
- Regex: `^[0-9]+$`
- If not digits-only, log as `Non-numeric input` and skip.

## Operational Guardrails
- Do not paste full list at once; always loop one value at a time.
- Keep waits configurable for different terminal response speeds.
- Run in attended mode first until stable in production environment.
