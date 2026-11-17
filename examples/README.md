# Inventory walkthrough

These sample files use fictional stock counts. Use a test Google spreadsheet that you can edit.

1. Open [inventory-source.csv](inventory-source.csv) in the app or save it as `.xlsx` in Excel first. Import [inventory-destination.csv](inventory-destination.csv) into a new Google Sheets tab and open its URL in the app.
2. In **Match by key** (*Cập nhật theo mã*), set the source range to `A1:C5` and the destination range to `A1:C6`.
3. Pair `A → A` and mark it **Key** (*Khóa*). Pair `B → B` and mark it **Write as number** (*Ghi thành số*). Enable **Add new keys** (*Thêm mã mới vào cuối vùng đích*).
4. Click **Preview** (*Xem trước*). Check for a quantity of `24` in `B4`, `12` in `B2`, and new item `VT-04` with a quantity of `0` in `A5:B5`. The notes in column C stay outside the mapping.
5. Click **Sync** (*Đồng bộ*). A second preview should show zero pending updates.

For a **Copy by position** test, create a separate blank tab, select `A1:C5` on the source side, and choose `A1` as the destination. Preview the write before applying it.
