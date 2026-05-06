# Batching ROI - DRIVES Standalone Flow Design

## Scope
This document defines a standalone DRIVES batching flow that is separate from the existing Interstate ROI mailbox automation, but stored in the same repository/workspace.

## DRIVES Navigation from Home/Main Menu
Use this exact menu path:

1. Option 3 - Law Enforcement Menu
2. Option 5 - Licence Menu
3. Option 14 - Certificates Menu
4. Option 7 - Batch Traffic Record Request
5. Enter

## Batch Request Selection Sequence (Confirmed)
Use this exact key sequence for the selection prompts:

1. Send keys: Y
2. Send keys: N
3. Send keys: Y
4. Send keys: {Enter}
5. Send keys: {Tab}
6. Send keys: S
7. Send keys: {F5}

After this, proceed to customer/licence number entry.

## Power Automate Desktop Keystroke Outline

```text
Focus DRIVES/JVT window

Send keys: 3
Send keys: {Enter}
Wait

Send keys: 5
Send keys: {Enter}
Wait

Send keys: 14
Send keys: {Enter}
Wait

Send keys: 7
Send keys: {Enter}
Wait

Send keys: Y
Send keys: N
Send keys: Y
Send keys: {Enter}
Wait

Send keys: {Tab}
Send keys: S
Send keys: {F5}
Wait

Start entering customer/licence numbers
```

## Batch Input Behaviour Rules

### 8-digit number
- Type number only.
- DRIVES should auto-advance to the next field.
- No tab required.

### 4-digit or 7-digit number
- Type number.
- Send keys: {Tab}

### Invalid length (not 4, 7, or 8)
- Do not enter into DRIVES.
- Add value to exception log.
- Continue to next number (or stop based on run setting).

## Input Loop Logic

```text
For each number in BatchList:
    Type number

    If length(number) = 8:
        // auto-advance expected in DRIVES
        Do nothing

    Else if length(number) = 4 or 7:
        Send keys: {Tab}

    Else:
        Log exception (invalid length)
        Continue

    Wait briefly
```

## Implementation Notes
- Keep this as a separate desktop flow from Interstate ROI email processing.
- Parameterize short wait durations so timing can be tuned per environment.
- Add run logging for:
  - Navigation start/end
  - Count of numbers entered
  - Count of invalid lengths
  - Any screen/cursor mismatch failures
