# Power Automate Desktop Build Checklist

## Flow Name
`DRIVES_Batching_ROI_Standalone`

## Goal
Navigate DRIVES to Batch Traffic Record Request, apply the confirmed selection sequence, and enter customer/licence numbers with length-based cursor behavior.

## Inputs
- `BatchList` (List): values to enter (string list).
- `DrivesWindowTitle` (Text): terminal window title to activate.
- `BatchFilePath` (Text, optional): full local path to a text file containing one number per line.

## Config Variables
Create these at the top of the flow:
- `WaitShortMs` = `300`
- `WaitMediumMs` = `600`
- `WaitLongMs` = `1000`
- `EnteredCount` = `0`
- `InvalidCount` = `0`
- `ExceptionLogPath` = `%CurrentUserDocuments%\\DRIVES_Batching_Exceptions_%FormatDateTime(CurrentDateTime, "yyyyMMdd_HHmmss")%.csv`

## Main Flow Steps

1. `Set variable` -> initialize config and counters.
2. `Write text to file` -> create CSV header:
   - Path: `ExceptionLogPath`
   - Text: `Timestamp,Value,Reason`
   - Append new line: `Yes`
3. `Activate window` -> title contains `DrivesWindowTitle`.
4. `Wait` -> `WaitMediumMs`.

### Build BatchList from file (if supplied by cloud flow)
5. `If` `BatchFilePath` is not empty and file exists:
6. `Read text from file` -> `BatchFileText`
7. `Split text` by new line -> `BatchList`
8. `End`

### Navigate Menu 3 > 5 > 14 > 7
9. `Send keys` -> `3`
10. `Send keys` -> `{Enter}`
11. `Wait` -> `WaitShortMs`
12. `Send keys` -> `5`
13. `Send keys` -> `{Enter}`
14. `Wait` -> `WaitShortMs`
15. `Send keys` -> `14`
16. `Send keys` -> `{Enter}`
17. `Wait` -> `WaitShortMs`
18. `Send keys` -> `7`
19. `Send keys` -> `{Enter}`
20. `Wait` -> `WaitMediumMs`

### Confirmed Selection Sequence
21. `Send keys` -> `Y`
22. `Send keys` -> `N`
23. `Send keys` -> `Y`
24. `Send keys` -> `{Enter}`
25. `Wait` -> `WaitMediumMs`
26. `Send keys` -> `{Tab}`
27. `Send keys` -> `S`
28. `Send keys` -> `{F5}`
29. `Wait` -> `WaitLongMs`

### Batch Entry Loop
30. `For each` item `CurrentNumber` in `BatchList`
31. `Trim text` on `CurrentNumber` -> `CurrentNumberTrimmed`
32. `If` `CurrentNumberTrimmed` is empty: `Continue loop`
33. `Get subtext` or `Replace text using regex` to keep digits only (optional hardening)
34. `Get length of text` -> `NumberLength`
35. `If NumberLength = 8`
36. `Send keys` -> `%CurrentNumberTrimmed%`
37. `Set variable` `EnteredCount = EnteredCount + 1`
38. `Wait` -> `WaitShortMs`
39. `Else If NumberLength = 4 OR NumberLength = 7`
40. `Send keys` -> `%CurrentNumberTrimmed%`
41. `Send keys` -> `{Tab}`
42. `Set variable` `EnteredCount = EnteredCount + 1`
43. `Wait` -> `WaitShortMs`
44. `Else`
45. `Set variable` `InvalidCount = InvalidCount + 1`
46. `Write text to file` append line:
   - `%FormatDateTime(CurrentDateTime, "yyyy-MM-dd HH:mm:ss")%,%CurrentNumberTrimmed%,Invalid length (%NumberLength%)`
47. `End`
48. `End loop`

### Run Summary
49. `Display message` or `Log message`:
   - `Entered: %EnteredCount% | Invalid: %InvalidCount% | Exceptions: %ExceptionLogPath%`

## Error Handling
- Add `On block error` around navigation and batch entry sections.
- In error branch:
  - append error row to `ExceptionLogPath`
  - capture screenshot (optional)
  - stop flow with clear message.

## Recommended First Test
1. Use a short list: `12345678`, `1234`, `1234567`, `99`, `123456789`.
2. Confirm behavior:
- 8-digit auto-advances.
- 4/7-digit advances only after tab.
- invalid values are logged, not entered.
3. Tune waits only if host latency requires it.
