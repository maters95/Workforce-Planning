# Cloud Flow Design - Mailbox to BatchList for DRIVES

## Objective
Read messages from ROI.InterstateRequests Inbox/(Batched Yesterday)Awaiting PDF, extract licence/customer numbers from subject lines, build a deduplicated list, and hand it to the desktop flow that enters DRIVES.

## Recommended Architecture
1. Cloud flow (scheduled or manual):
- Reads mailbox items.
- Extracts likely number tokens from subjects.
- Applies validation (length 4/7/8 only, digits only).
- Deduplicates values.
- Writes the final list to a text file (one number per line).
2. PAD flow:
- Reads text file into `BatchList`.
- Runs DRIVES navigation and entry logic.

## Cloud Flow Actions (Power Automate Cloud)

1. Trigger
- `Recurrence` (or `Manually trigger a flow`).

2. Resolve target folder ID
- `Send an HTTP request (Microsoft Graph)`
- Method: `GET`
- URI:

```text
https://graph.microsoft.com/v1.0/users/ROI.InterstateRequests@transport.nsw.gov.au/mailFolders/inbox/childFolders?$select=id,displayName&$top=100
```

- `Filter array` where `displayName` equals `(Batched Yesterday)Awaiting PDF`.

3. List messages
- `Compose` -> `FolderMessagesUri`

```text
https://graph.microsoft.com/v1.0/users/ROI.InterstateRequests@transport.nsw.gov.au/mailFolders/@{first(body('Filter_Awaiting_PDF_Folder'))?['id']}/messages?$select=id,subject,receivedDateTime&$top=999
```

- `HTTP` GET to `outputs('FolderMessagesUri')`.

4. Initialize variables
- `varNumbersRaw` (Array) = `[]`
- `varNumbersValid` (Array) = `[]`
- `varInvalidRows` (Array) = `[]`

5. For each message
- Input: `body('HTTP_List_Messages')?['value']`

6. Normalize subject
- `Compose SubjectLower`:

```text
toLower(trim(coalesce(items('Apply_to_each_Message')?['subject'], '')))
```

- `Compose SubjectClean` (replace punctuation with spaces):

```text
replace(replace(replace(replace(replace(replace(replace(replace(replace(outputs('SubjectLower'),'(',' '),')',' '),',',' '),'-',' '),'/',' '),':',' '),'.',' '),';',' '),'#',' ')
```

- `Compose Tokens`:

```text
split(outputs('SubjectClean'),' ')
```

7. For each token in Tokens
- `Compose TokenTrim`:

```text
trim(item())
```

- Skip if empty.

- Digits-only check expression:

```text
equals(trim(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(outputs('TokenTrim'),'0',''),'1',''),'2',''),'3',''),'4',''),'5',''),'6',''),'7',''),'8',''),'9','')), '')
```

- Length check expression:

```text
or(equals(length(outputs('TokenTrim')),4),equals(length(outputs('TokenTrim')),7),equals(length(outputs('TokenTrim')),8))
```

- If both true:
  - `Append to array variable` -> `varNumbersRaw` value `outputs('TokenTrim')`
- Else if digits-only true but wrong length:
  - Append object to `varInvalidRows`:

```json
{
  "messageId": "@{items('Apply_to_each_Message')?['id']}",
  "subject": "@{items('Apply_to_each_Message')?['subject']}",
  "token": "@{outputs('TokenTrim')}",
  "reason": "Invalid length"
}
```

8. Deduplicate
- `Compose UniqueNumbers`:

```text
union(variables('varNumbersRaw'), variables('varNumbersRaw'))
```

9. Convert to newline text for PAD
- `Compose BatchListText`:

```text
join(outputs('UniqueNumbers'), decodeUriComponent('%0A'))
```

10. Write handoff file
- Create file in SharePoint/OneDrive path used by machine runtime, e.g.:

```text
DRIVES/BatchInput/BatchList.txt
```

- File content = `outputs('BatchListText')`

11. Optional invalid log file
- Create `InvalidTokens.json` with content `string(variables('varInvalidRows'))`.

12. Run desktop flow
- `Run a flow built with Power Automate for desktop`
- Input parameter `BatchFilePath` pointing to synced local path.

## Subject Parsing Behavior
- This approach captures numeric tokens anywhere in the subject.
- It accepts only digits and only lengths 4, 7, 8.
- It ignores text tokens and non-numeric IDs.

## Handoff Contract to PAD
- `BatchList.txt` must contain one number per line.
- Example:

```text
12345678
1234
1234567
```

## Reliability Notes
- Keep extraction and DRIVES entry in separate flows for easier troubleshooting.
- Save both the produced list and invalid token log each run for auditability.
- Add alert if `length(outputs('UniqueNumbers')) = 0` to avoid opening DRIVES with no data.