<h1 align="center">Excel Sheets Sync</h1>

<p align="center">Preview and send changes from a local Excel file to Google Sheets.</p>

<p align="center"><kbd>v0.2.1</kbd> &nbsp; <kbd>Windows x64</kbd> &nbsp; <kbd>Vietnamese interface</kbd></p>

![Excel Sheets Sync showing two inventory tables and a preview of four planned cell updates](docs/demo.png)

## Get started

1. Extract the Windows release ZIP and open `ExcelSheetsSync.exe` from the extracted folder. Keep the runtime files and folders beside the executable. The portable build includes Electron and reads saved workbooks directly.
2. Choose **Open demo data** (*Mở dữ liệu mẫu*) → **Preview** (*Xem trước*) → **Sync** (*Đồng bộ*). This updates the in-app sample. Opening the demo again resets it.
3. To try a real connection, complete [Google sign-in setup](#google-sign-in-setup), then follow the [inventory walkthrough](examples/README.md) with the included fictional data.

**Using a managed Windows PC?** The v0.2.1 executable is unsigned. SmartScreen or your organization's application policy may restrict running it. [Microsoft SmartScreen guidance](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)

## Sync a workbook

Drag a saved workbook into the **Excel source** pane, paste a Google Sheets URL into the **Google Sheets destination** pane, and click **Open Sheets** (*Mở Sheets*). Select the tabs and ranges you want to use.

| Mode | How it works |
| --- | --- |
| **Copy by position** (*Chép theo vị trí*) | Select a source range and the top-left destination cell. The app places values at the corresponding offsets. |
| **Match by key** (*Cập nhật theo mã*) | Include one header row at the top of both ranges. Pair the columns by dragging their letters or clicking the source and destination letters in order. Mark one or more pairs as **Key** (*Khóa*) to match rows even when the order differs. |

Review **Preview**, then click **Sync** to apply the proposed changes. Give the mapping a name and choose **Save mapping** (*Lưu cách ghép*) to reuse it on the same Windows profile.

- **Add new keys** (*Thêm mã mới vào cuối vùng đích*) fills available rows at the end of the selected destination range. Include blank rows in that range and map every column you want to populate.
- **Write as number** (*Ghi thành số*) converts Vietnamese-formatted text such as `22.000.000 ₫` into a numeric value in a mapped data column. Key columns retain their identifiers.
- **Expand** (*Phóng to*) opens either pane at full width with the other pane in a small corner preview. Click the small preview to switch sides. Use **Split** (*Chia đôi*) or `Esc` to return to the default view.
- Use the scrollbars or a touchpad to move in both directions. Use `Shift` + wheel for horizontal scrolling, or hold `Space` and drag to pan. Drag a selection to the grid edge to extend it beyond the visible area.

## Google sign-in setup

Create one **Desktop app OAuth client** for the app. Each tester uses their own Google account and the file permissions granted to that account.

1. Select or create a Google Cloud project and enable the [Google Sheets API](https://console.cloud.google.com/apis/library/sheets.googleapis.com).
2. Set up [Google Auth Platform](https://console.cloud.google.com/auth/overview). For an **External** app in **Testing**, add each tester under **Audience → Test users**. Google authorizations for the requested Sheets scope expire after seven days in that mode. [Google OAuth testing guidance](https://support.google.com/cloud/answer/15549945?hl=en)
3. Open [Clients](https://console.cloud.google.com/auth/clients), create an OAuth client with the **Desktop app** type, and download its JSON file. [Google credential setup](https://developers.google.com/workspace/guides/create-credentials#desktop-app)
4. In the app, click **Sign in with Google** (*Đăng nhập Google*) → **Choose JSON and sign in** (*Chọn JSON và đăng nhập*). Alternatively, place the desktop client JSON beside the portable executable as `google-oauth.json` before launch. For a source checkout, place it in the project root.
5. Sign in with an account that can read the destination. To sync, the account needs **Editor** access and permission to edit the selected cells.

Keep OAuth JSON files, tokens, and working spreadsheets outside version control. Windows stores the app's encrypted Google session, saved mappings and sync history under `%APPDATA%\Excel Sheets Sync`. The app sends the selected updates directly to the Google Sheets API. Saved mappings and sign-in sessions are local to the Windows profile.

> **403 `access_denied` while signing in?** Check that the signed-in email is listed under **Test users** in the same project as the imported OAuth client. For wider distribution, review Google's requirements for the app's [sensitive Sheets scope](https://developers.google.com/workspace/sheets/api/scopes).

## Build and test

Use Windows x64 with Node.js **22.12+** and npm.

```powershell
npm.cmd ci
npm.cmd start
npm.cmd test
npm.cmd run test:desktop
npm.cmd run package:win
```

The desktop smoke test uses generated data, a temporary app profile, and simulated Google responses. The release build is written to `release/ExcelSheetsSync-win32-x64`. To distribute it, ZIP the entire output folder. The `release/` directory is ignored by Git.

## Scope of v0.2.1

- Manual, one-way sync into an existing Google Sheets tab. Sources can be `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`, or `.tsv`, up to 35 MiB and 30,000 cells per selected range.
- Sync writes cell values, including cached Excel formula results. It skips blank source cells, preserves zero, flags duplicate keys, and blocks writes into destination formula cells, formula results, or merged cells. Google Sheets applies the destination's existing number format.
- Preview shows basic formatting and merged cells. Charts, embedded images, and frozen panes are outside the preview's scope.
- The app rechecks the plan before writing. For simultaneous use, assign distinct destination rows or ranges to each editor to avoid races on shared cells.
