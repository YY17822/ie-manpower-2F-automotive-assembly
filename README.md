# IE 人力需求程式（2F 組裝）

IE 部門每月月底計算未來六個月人力需求的小工具。一步一步做，共 9 步，第 9 步產生總表。

名詞：**行** = 直的（A、B、C…），**列** = 橫的（1、2、3…）。

## 怎麼開啟

- **最簡單**：點兩下 `index.html`，用 Chrome 或 Edge 開啟。
- **要安裝成 App、離線使用**：點兩下 `啟動小程式.command`，瀏覽器會開啟 `http://localhost:3004`。用完把跳出來的終端機視窗關掉。
- **用 VS Code Go Live**：在 VS Code 開啟專案資料夾，按右下角「Go Live」，瀏覽器會開啟 `http://127.0.0.1:3004`（port 在 `.vscode/settings.json`）。

第一次使用需要連上網路，小程式會載入 ExcelJS（讀寫 Excel 的元件）。
要在不能連外部網站的電腦上用，請看 `vendor/README.txt`。

## 第 1 步：篩選 MPS 報表（好幾份合成一份，依類別分段）

1. 上傳 MPS 的 Excel，可以一次選好幾份（只支援 `.xlsx`；舊的 `.xls` 請先在 Excel 另存成 `.xlsx`）。之後還可以再加入、移除或調整順序。依上傳順序叫報表 1、報表 2…。
   每份報表的**類別**（取像、車用、光機系統）依檔名預設：有 `TYO` 是光機系統、有 `TAI` 是車用、有 `Image` 是取像（`TYO`、`TAI` 要大寫，`Image` 不分大小寫）；都沒有的話在清單裡自己選，沒選完不能匯出。
   報表**依類別排順序**（取像 → 車用 → 光機系統 → 還沒選的），依序叫報表 1、報表 2…；同一類有好幾份時照上傳順序。改了類別會重新排。
   每份報表的條件照**位置**記住：「類別 + 同類別的第幾份」（例如報表 1 是取像的第 1 份），和上傳順序、一次傳幾份都無關。報表排到哪個位置，就帶入那個位置上次的條件；在那個位置改的條件馬上記在那個位置。移除報表時，記住的條件不會刪掉，後面遞補上來的報表會帶入新位置的條件。還沒選類別的報表不帶入任何條件；這時先輸入的條件，選了類別後會保留下來。
   只有同一類有好幾份時才能「往上移」（換這兩份的順序，也就是換帶入的條件）。
   各報表的預覽會畫出整份報表（最多 3000 列），可以上下、左右捲動。
   工作表預設：報表 1 選「MPS」、報表 2 選「TAI MPS」、報表 3 選「3DP MPS」（名稱不分大小寫、不管空白）；檔案裡沒有這個名稱的工作表就用第一個。自己選過工作表的報表不會再自動換。
2. 每份報表各自輸入：
   - 篩選方式（每份報表各自選，可以混用）：
     - 「A 行等於」（原本的方式）：A 行等於指定值的列才會留下，從第 5 列開始比對。可以輸入多個值，例如 `1, 3`；也可以點報表 A 行出現過的值。
     - 「指定行 等於／不等於」：自己指定一行（例如 `D`），選「等於」或「不等於」，再輸入值（多個用逗號分開，也可以點那一行出現過的值）。那一行**空白的列一律不要**。選「不等於」又沒填值，就是只拿掉那一行空白的列。
     - 「指定某一行，模糊比對（含有某些字）」：指定一行，選「選取」或「排除」，在輸入框填要比對的字（預設 3 個輸入框，可以「＋ 再加一個輸入框」，也可以刪除，至少留 1 個）。儲存格的字**含有**任一個輸入框的內容就算符合，英文大小寫不分、全形半形視為相同。「選取」留下符合的列，「排除」留下都不符合的列；那一行**空白的列一律不留**。輸入框下面會逐列列出篩選出來的列（含有的字用螢光筆標出，最多列 300 列），並註明排除幾列、空白幾列不留；原始報表預覽裡不會留下的列會變淡。點「報表裡 X 行的值」的按鈕會填進第一個空的輸入框，已經有這個值就清掉。
     - 「依區塊（TTL）」：指定 TTL 寫在哪一行（預設 A 行）。那一行寫著「XXXX TTL」的列是一個區塊的結尾，它上面、到前一個 TTL 列為止的列就是「XXXX」這個區塊（例如「3DP MPS TTL」上面那些料號是「3DP MPS」）。畫面會列出找到的區塊（幾列、第幾列到第幾列、TTL 在第幾列），勾選要的區塊；TTL 那一列和那一行空白的列不留。最後一個 TTL 下面如果還有列，會列成「（最後一個 TTL 之後）」。勾的區塊名稱會記住，下次同名的區塊自動勾好。
   - 要保留的行：例如 `B, C, U, Z, AE`，連續的行可寫 `U:Z`；也可以直接點預覽表格上方的行號。報表 2 之後有「和報表 1 保留一樣的行」可以一鍵帶入。
   - 各報表保留的行位置可以不同，但**行數和順序要一樣**：每份報表保留的第 n 個行，都放在新檔的第 n 行。行數不一樣時不能匯出。
3. 挑出 RMA：從上面各報表留下來的**所有列**裡，用一組條件挑出 RMA（所有報表共用）：
   - 比對的行：選新檔的某一行（各報表保留的第 n 個行都在新檔第 n 行）。不選就不挑 RMA。
   - 篩選方式二選一：
     - 「指定某一行，模糊比對（含有某些字）」：那一格的字含有任一個輸入框的內容就是 RMA（英文大小寫、全形半形不分；輸入框預設 3 個，可以加、可以刪）。
       下面逐列列出符合的列（含有的字用螢光筆標出），每列前面有勾勾，不是 RMA 的**取消勾選**。取消勾選只在這次有效，不會記住。
     - 「指定某一行，等於某些值」：那一格等於任一個值就是 RMA（多個值用逗號分開）。下面一樣列出符合的列（沒有勾勾）。
   - 點「留下來的列裡，新檔 X 行的值」的按鈕可以直接填入。那一行空白的列不會是 RMA。
   - 挑出來的列會從【取像】【車用】【光機系統】移走，放到最下面的【RMA】。
   - 「RMA 區預留 N 列」（預設 6，會記住）：【RMA】區是挑出的列放最上面，下面再接 N 列空白列（細格線）。
     列數 = 挑出的列數 + N（沒挑出時就是 N 列空白）。設 0 就是不預留。
     空白列的 CT.、工作中心、需求工時都先放好公式（見第 4、5 步），產出 Excel 後直接填料號和數量就會自動算；
     RMA 的 Total、RMA總需求工時、總表的 MPS(PCS) RMA 和需求工時-重工的範圍都包含這些列，不用改公式。
     RMA 區有幾列記在 Excel 檔的「關鍵字」屬性裡（`ie-manpower:rmaRows=N`），後面的步驟、中途匯出再上傳的檔案都靠它找到預留列。
4. 選擇 Status 在哪一行（區分料號是量產還是試產，後面的總表會用到）。從報表 1 保留的行裡選，表頭有「Status」的話會自動選好；畫面會列出其他報表在同一個位置的行和表頭，表頭和報表 1 不一樣時會提醒。
5. 按「匯出 Excel」，得到 `報表1檔名_等N份_篩選.xlsx`（只有一份時是 `原檔名_篩選.xlsx`）。

匯出的檔案：

- 第 1–4 列表頭用報表 1 的，原樣保留，包含格式、行寬、列高、凍結窗格。
- 表頭下面依類別分段，順序固定是 取像 → 車用 → 光機系統（和上傳順序無關）：

  | 位置 | 內容 |
  | --- | --- |
  | 表頭下一列 | 【取像】類別標題 |
  | | 取像的 MPS（取像的報表照上傳順序往下接） |
  | 空 3 列 | |
  | | 【車用】類別標題 |
  | | 車用的 MPS |
  | 空 1 列 | |
  | 3 列 | 第 2 步的匯總表「取像＋車用」（Total、量產、試產；第 1 步先留空） |
  | 空 6 列 | |
  | | 【光機系統】類別標題 |
  | | 光機系統的 MPS |
  | 空 1 列 | |
  | 3 列 | 第 2 步的匯總表「光機系統」 |
  | 空 3 列 | |
  | | 【RMA】類別標題（底色 `#FFCAD4`） |
  | | 挑出來的 RMA（照它們原本在 取像 → 車用 → 光機系統 裡的順序） |
  | 空 1 列 | |
  | 1 列 | 第 2 步的「RMA」Total（只有 Total 一列，不分量產、試產） |

- 類別標題那一列跨所有保留的行合併，粗體置中，底色 `#E9C46A`。每一段上面一定有它的類別標題。
- 這個月沒有檔案的類別整段不放（標題、空列、匯總表都不放）；例如沒有取像時，【車用】直接接在表頭下面。一類的列全部挑去 RMA 時，那一類也整段不放。
- 【RMA】一定會放：沒有挑出 RMA 時，標題和 Total 一樣會放，只是沒有項目（Total 是 0）。它接在最後一張匯總表下面空 3 列。
- 同一類有好幾份報表時照上傳順序往下接，中間不畫線（以前的紫色分隔線已經拿掉）。原報表帶過來的黑色粗框線會拿掉，各段裡面只留細線。
- 同一個料號出現在不同報表，會分開保留成好幾列。畫面上會列出各段的標題在第幾列、資料在第幾列到第幾列，以及匯總表會放在哪裡。
- 後面的步驟用類別標題（第 1 行的字剛好是「取像」「車用」「光機系統」）找出各段的範圍（`ExcelIO.mpsLayout`），只對資料列做事；標題列、空列、匯總表都會跳過。沒有類別標題的舊檔案，整塊資料當作一段。
- 合併儲存格如果只保留了其中幾行，會縮成只合併那幾行；大標題如果在沒保留的行，會移到左上角。
- 公式會換成計算結果的數值（篩掉列和行之後，原本的公式參照會錯位）。
- 所有儲存格的填滿色彩都會拿掉（變透明），包括表頭。字型、框線、數字格式照舊。
- 每份報表的類別、條件、Status 的位置會記住（依報表順序），下個月照同樣順序上傳就會自動帶入。
- Status 在新檔的哪一行，會記在 `IEApp.state.mpsFiltered.statusCol`（數字）和第 1 步的設定 `statusCol`（英文字母），給後面的步驟用。

報表版面不同時（表頭不是 4 列，或選取條件不在 A 行），在第 1 項下面的「報表版面不一樣時」調整，所有報表共用這組設定。

## 第 2 步：匯總

算出 MPS 各月的預計總產量（Total），再用 Status 分成量產（M）和試產（R）的合計。

1. 第 1 步的 Excel：剛做完第 1 步的話會直接沿用；也可以上傳之前第 1 步匯出的檔案。
2. 確認 Status 在哪一行（預設是第 1 步指定的那一行）、當月 WIP 在哪一行、各月產量在哪幾行（第一個是當月）。
3. 按「匯出 Excel」，得到 `…_篩選_匯總.xlsx`。

匯總表有兩張，各放在那一組最後一列資料之後空一列：

- **取像＋車用**：取像和車用兩段一起算，放在車用（沒有車用就是取像）的下面。
- **光機系統**：只算光機系統，放在光機系統的下面。
- **RMA**：只有 Total 一列，放在【RMA】最後一列下面空一列（沒有項目時在標題下面空一列，填 0）；旁邊一樣註明「RMA」。
- 沒有類別標題的舊檔案：整塊資料一張，放在最後一列資料之後空一列。

結果（只有一段時，以 Status 在 C 行、WIP 在 D 行、當月在 E 行、資料在第 5 到 74 列為例）：

| 列 | 當月（要加 WIP） | 其他月份 |
| --- | --- | --- |
| Total | `=SUM(D5:D74)+SUM(E5:E74)` | `=SUM(F5:F74)` |
| 量產 | `=SUMIF($C$5:$C$74,"<>R",D5:D74)+SUMIF($C$5:$C$74,"<>R",E5:E74)` | `=SUMIF($C$5:$C$74,"<>R",F5:F74)` |
| 試產 | `=SUMIF($C$5:$C$74,"R",D5:D74)+SUMIF($C$5:$C$74,"R",E5:E74)` | `=SUMIF($C$5:$C$74,"R",F5:F74)` |

- 取像＋車用有兩段時，每個範圍都寫兩段，中間的標題列、空列不會算進去，例如 `=SUM(E6:E40,E45:E80)`、`=SUMIF($C$6:$C$40,"M",E6:E40)+SUMIF($C$45:$C$80,"M",E45:E80)`。
- 位置：每張三列，對齊各月的行；文字標籤在第一個月左邊那一行（通常就是 WIP 那一行）。標籤再左邊一行跨三列合併、粗體，註明是「取像＋車用」還是「光機系統」（那一行是料號的行時不放）。
- 當月一律先加 WIP。
- Status 是 R 算試產，**其他的（包括 M、其他字、空白）都算量產**（不分大小寫），所以量產 + 試產 = Total（和第 6 步同一個規則）。Status 那一行完全沒有 R 時會提醒。
- Total 那一列是淺綠色底（`#B5E48C`）；標籤加上各月圍成一塊，外框粗線、裡面細格線。
- 不再標示前三大（舊版的淺橘色條件式格式，重做時會拿掉）。
- 料號固定看第 1 行（用來判斷哪些列是資料）。
- 這兩張和第 6 步的兩張需求工時匯總在同一列上，一個在左邊（產量）、一個在右邊（工時），互不影響。

## 第 3 步：匯入 P003 標準工時

把 P003 報表（公司所有產品型號的標準工時）放進第 1 步完成的 Excel，成為同一個檔案裡名為 `P003` 的工作表。

1. 第 2 步的 Excel：剛做完第 2 步的話會直接沿用，不用先匯出；也可以上傳之前第 2 步匯出的檔案。跳過第 2 步的話，會沿用第 1 步篩選好的結果。
2. 上傳 P003 報表（`.xlsx`）。檔案裡有多個工作表時可以選要用哪一個。
3. 按「匯出 Excel」，得到 `…_匯總_含P003.xlsx`，裡面有原本的 MPS 工作表和 `P003` 工作表。

P003 工作表是整張照抄：值、格式、合併儲存格、行寬、列高、凍結窗格、自動篩選都會帶過去，公式換成數值。
上傳的檔案裡如果已經有 `P003` 工作表，會換成這次上傳的。

## 第 4 步：帶入標準工時 CT.

拿 MPS 工作表的料號到 P003 工作表找同一個料號，把那一列的 Labor ÷ Base Qty 帶回 MPS。
P003 的 Labor 是做 Base Qty 個的工時（Base Qty 有 1 也有 100），除過之後才是單件工時。

1. 第 3 步的 Excel：剛做完第 3 步的話會直接沿用；也可以上傳之前第 3 步匯出的檔案（要有 MPS 和 P003 兩個工作表）。
2. 選 MPS 的料號在哪一行。
3. 選 P003 的料號、Labor、Base Qty、Work Center 各在哪一行（表頭有 `Work Center` 的那一行會自動選好；不想帶入就選「不帶入工作中心」）。
4. 按「匯出 Excel」，得到 `…_含P003_CT.xlsx`。

結果：

- CT. 放在 MPS 工作表「原本最後一行之後空一行」的位置，表頭是 `CT.`。
- 填入的是公式，例如 `='P003'!M5/'P003'!K5`（Labor ÷ Base Qty），在 Excel 點儲存格就能看到數值是從 P003 哪兩格算出來的。
- 同一個料號在 P003 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）；畫面上會列出各列算出來的 CT. 不一樣的料號。這一點之後可能再調整。
- Base Qty 是空白或 0 的料號，Excel 會顯示 `#DIV/0!`，畫面上會列出是哪些料號。
- P003 找不到的料號，CT. 留空白並標上淺紅色底，畫面上會列出是哪些料號。
- 料號比對不分大小寫、會去掉頭尾空白；`0012` 和 `12` 視為不同。
- 對已經有 `CT.` 的檔案重做這一步，會覆蓋原本那一行，不會再多一行。
- CT. 右邊緊接著一行 `工作中心`：**公式**參照 P003 同一列的 Work Center，例如 `='P003'!N5`，點儲存格看得到是哪一格（P003 那一格是空白時寫 `=IF('P003'!N5="","",'P003'!N5)`，才不會顯示 0）。整行置中。P003 找不到的料號留空白、淺紅色底。重做時一樣覆蓋原本那一行。
- 因為多了工作中心這一行，第 5 步的需求工時會往右移一行（一樣在最後一行之後空一行）。
- RMA 區的空白預留列：CT. 放查表公式，例如 `=IFERROR(INDEX('P003'!M:M,MATCH(A86,'P003'!A:A,0))/INDEX('P003'!K:K,MATCH(A86,'P003'!A:A,0)),"")`；工作中心是 `=IF(A86="","",IFERROR(INDEX('P003'!N:N,MATCH(A86,'P003'!A:A,0))&"",""))`。填了料號就自動抓，料號空白或 P003 找不到時顯示空白。公式裡的行照這一步選的料號、Labor、Base Qty、Work Center。

## 第 5 步：計算需求工時

把每個料號各月的計畫產量乘上該列的 CT.，算出各月的需求工時。

前提是第 1 步保留的行照這個順序：前三行是料號、型號、生產方式，第 4 行是當月 WIP，
第 5 到第 10 行是當月和往後五個月的計畫產量。順序不同時可以在畫面上改。

1. 第 4 步的 Excel：剛做完第 4 步的話會直接沿用；也可以上傳之前第 4 步匯出的檔案。
2. 確認當月 WIP 在哪一行、各月產量在哪幾行（預設是第 4 行和第 5 到 10 行）。這組設定和第 2 步共用，在其中一步改了，另一步也會跟著變。CT. 會自動找表頭是 `CT.` 的那一行。
3. 各月的表頭：當月從下拉選單選（`Jan.`～`Dec.`），再按最後面的「自動帶出其餘月份」，其餘月份會依序填好（`Dec.` 之後接 `Jan.`），填完還可以自己改。程式不會去 MPS 的表頭找。
4. 按「匯出 Excel」，得到 `…_CT_需求工時.xlsx`。

結果：

- 從 MPS 工作表「原本最後一行之後空一行」開始，一個月一行。
- 每格都是公式：當月是 `=(D5+E5)*L5`（WIP + 當月產量，再乘 CT.），其餘各月是 `=F5*L5`。
- 表頭：第 2 列是跨所有月份置中的「需求工時」；第 3、4 列合併，寫你輸入的月份。
- 數值只顯示整數。儲存格裡仍保留小數，加總用的是完整的數值，所以合計可能和「把畫面上的整數加起來」差 1。
- 沒有 CT. 的料號一樣會填公式，結果是 0，並標上淺紅色底；畫面上會列出是哪些料號。
- 數值的單位和 CT. 相同（P003 的 Labor 單位是分鐘，所以這裡是分鐘），不換算。
- 對已經有「需求工時」的檔案重做這一步，會覆蓋原本的位置。
- RMA 區的空白預留列也先放好公式，料號空白時顯示空白：當月 `=IF($A86="","",(D86+E86)*L86)`，其餘月 `=IF($A86="","",F86*L86)`。

## 第 6 步：需求工時匯總

在需求工時下面加匯總，每格都是公式。和第 2 步一樣有三張（取像＋車用、光機系統、RMA），放在和第 2 步那三張同一列，只是在需求工時那幾行底下；取像＋車用的公式一樣寫兩段範圍。RMA 那張只有一列「RMA總需求工時」（淺藍色底）。


| 列 | 公式（以 Status 在 C 行、需求工時在 N 行、資料在第 5 到 74 列為例） |
| --- | --- |
| Production 總需求工時 | `=SUM(N5:N74)` |
| 量產總需求工時 | `=SUMIF($C$5:$C$74,"<>R",N5:N74)` |
| 試產總需求工時 | `=SUMIF($C$5:$C$74,"R",N5:N74)` |

- Status 在哪一行：和第 1 步、第 2 步共用同一個設定（第 1 步指定，在第 2 步或第 6 步改了，三步都跟著變）。
- Status 是 R 算試產，**其他的（包括 M、其他字、空白）都算量產**（不分大小寫），所以量產 + 試產 = Production。（和第 2 步的產量匯總同一個規則。）
- 文字標籤放在需求工時左邊那一行（原本空著的那一行），置中；那一行會加寬到放得下標籤。
- 標籤加上各月圍成一塊：外框粗線、裡面細格線。「Production 總需求工時」那一列是淺藍色底（`#A2D2FF`）。
- 不再用顏色標示前三大（舊版的淺橘色條件式格式，重做時會拿掉）。
- 對已經有匯總的檔案重做這一步，會覆蓋原本那幾列。

## 第 7 步：匯入人力報表

（原本的第 7 步「RMA 需求工時」已經拿掉，之後改成在第 1 步挑選哪些列是 RMA。這一步和總表往前一號：檔名、共用設定的組名也跟著改成 `step7`、`step8`，以前存的設定要重新輸入一次，見下面「拿掉 RMA 之後的重新編號」。）

把人力報表（未來六個月的人力預計）放進前面做好的 Excel，成為同一個檔案裡名為 `人力` 的工作表，
並確認之後總表要抓的資料在哪裡。

1. 第 6 步的 Excel：剛做完第 6 步的話會直接沿用；也可以上傳之前第 6 步匯出的檔案。
2. 上傳人力報表（`.xlsx`）。檔案裡有多個工作表時可以選要用哪一個（有叫「人力」的會先選它）。
3. 確認六個項目各在第幾列：Online、Online（生產）、Online（重工）、Off line（直接）、Off line（間接）、TTL（On + Off直接）；
   以及當月（第一個月）在哪一行。六個月是從當月那一行開始往右連續六行。
4. 按「匯出 Excel」，得到 `…_人力.xlsx`。

- 選單裡每一列會顯示「第 N 列　這一列的文字」。程式會先依文字帶入（例如有 `Online` 和 `生產` 的那一列就是 Online（生產）），請對照預覽確認，不對就改。
- 當月那一行：選好的各列裡的數字剛好落在六個連續的行時才會自動帶入；報表連過去的月份也列出來時，程式不知道哪一行是當月，要自己選。
- 預覽表格裡螢光的格子，就是之後總表要抓的資料；下面的核對表列出六個項目 × 六個月讀到的數值，空白或不是數字的格子標成淺紅色。
- 兩個項目選了同一列、或有空白的格子，會提醒但不擋匯出；六個項目和當月都要選好才能匯出。
- 六個項目都選好時，各項目的列會記在這台電腦的瀏覽器裡；下個月列的文字對不上時，就用記住的列當預設值。
- `人力` 工作表是整張照抄（值、格式、合併儲存格、行寬、列高），公式換成數值。上傳的檔案裡如果已經有 `人力` 工作表，會換成這次上傳的。
- 這一步匯出的檔案裡 `人力` 還是一個工作表；到第 8 步匯出總表時，會整張併進「總表」工作表的 O 行起，不再另外留 `人力` 工作表。
- 確認的位置放在 `IEApp.state.withManpower`（`rows`：各項目的列號，`firstMonthCol`：當月的行號，`monthCount`：6），給之後的總表用。

## 第 8 步：總表

在同一個檔案裡新增「總表」工作表（B 到 J 行）。總表排在第一個工作表，打開檔案時顯示總表。
第 7 步匯入的人力報表整張併進總表的 **O 行起**（列號不變；值、格式、合併儲存格、行寬照抄，公式換成數值；同一列的列高取比較高的），不再另外留 `人力` 工作表。

| 列 | 內容 |
| --- | --- |
| 2 | 大標題（B 到 J 合併，預設「製二課 (車用/光機)」，畫面上可以改，會記住） |
| 4 | 法定加班上限46hr/人（紅底白字，B、C 合併） |
| 5 | 表頭：B、C 合併畫斜線；D 到 I 是六個月（參照 MPS「需求工時」的月份）；J 是備註（深藍底白字） |
| 6–8 | MPS(PCS) 量產、試產、RMA（淺藍底） |
| 9–11 | 需求工時-量產、-試產、-重工（min/月） |
| 12–15 | 平日工作天數、假日工作天數、工作時間、加班時間 |
| 16–19 | 綜合效率% - 量產、- 試產、- 重工、出勤率% |
| 20–23 | 需求人力：On line、新機種+重工、Off line(20%)、TTL（淺綠底） |
| 24–26 | 現有人力：On line、Off line、TTL（淺綠底） |
| 27 | 人力需求差異(現有-需求)（上面一條粗線） |
| 28 | Loading Rate%(含加班)（淺橘底） |

人力相關的字（第 20 到 28 列）是藍色（`#4472C4`）。

M1 寫「MPS資料來源」（粗體），M2 寫第 1 步上傳的所有 MPS 檔名，一個檔名一行（同一格裡自動換行，照報表 1、2、3… 的順序；第 2 列會加高到放得下）。檔名在第 1 步就記在 Excel 檔的「描述」屬性裡（檔案 → 資訊 → 摘要），中途匯出再上傳的檔案也帶得到；這次改版之前匯出的舊檔案沒有這個記錄，M1、M2 就留空白。

各列的來源和公式（以第一個月 D 行為例）：

- MPS(PCS) 量產、試產：MPS 工作表第 2 步匯總的「量產」「試產」那一列（當月已含 WIP）。匯總表有兩張時兩張相加，例如 `='MPS'!E17+'MPS'!E30`。
- MPS(PCS) RMA：第 2 步【RMA】那張的 Total；需求工時-重工：第 6 步「RMA總需求工時」。RMA 是第 1 步挑出來的。
- 需求工時-量產、-試產：第 6 步「量產總需求工時」「試產總需求工時」，兩張相加。
- 平日工作天數、工作時間、綜合效率%（量產／試產／重工，第一次預設 95%、85%、80%，之後用上次輸入的值）、出勤率%（預設 95%）：畫面上輸入。
- 假日工作天數 `=(46-D12*2)/8`；加班時間 `=2*60*80%`（每日加班小時 × 60 × 加班出勤人數 %）。
- 需求人力 On line `=(D9/D16+D10/D17+D11/D18)/(D12+D13)/(D14+D15)/D19`，顯示到小數點第一位。
- 需求人力 新機種+重工：畫面上輸入各新機種（名稱可以增減）和重工各月幾人，加總填公式，例如 `=2+3+1`；備註寫「各月人數 = 機種A + 機種B + 重工」。
- 需求人力 Off line(20%)：畫面上輸入各項目和人數（每個月一樣；預設領班、物料（含材製損）、文書、包裝，可以「＋ 新增項目」、刪除、改名稱），有填人數的項目加總填公式，例如 `=1+2+1+3`；備註寫各項目的人數。
- 需求人力 TTL `=D20+D21+D22`。
- 現有人力 On line、Off line：第 7 步人力報表的 Online、Off line（直接），參照總表 O 行起的位置，例如 `='總表'!P2`；TTL `=D24+D25`。
- 人力需求差異 `=D26-D23`；Loading Rate `=D23/D26`（需求人力 TTL ÷ 現有人力 TTL）。

第二張表（第 32 到 56 列）：和第一張一樣，紅色格子寫「YO加班上限78hr/人」，假日工作天數改用 78（畫面上可以改）；其他輸入和抓來的數值都參照第一張表。

匯出的檔名預設是 `YYMM_製二課 人力需求預估.xlsx`：`YY` 是年份後兩碼，`MM` 是首月（第 5 步選的當月，例如 2026 年 11 月 → `2611_製二課 人力需求預估.xlsx`）。年份取離今天最近的那一年（例如 12 月做隔年 1 月的，年份是隔年）。檔名顯示在「匯出 Excel」旁邊，可以自己改；改過之後就不會再自動換，按「用預設檔名」可以換回來。

## 檔案

```
index.html                  主畫面
js/config.js                專案代號、Supabase 網址和公開鑰匙（SUPABASE_PUBLIC_URL、SUPABASE_PUBLISHABLE_KEY）
js/cloud.js                 登入與共用設定（Supabase）
manifest.webmanifest        安裝成 App 用的資訊
sw.js                       離線快取
css/style.css               樣式
js/app.js                   外框：步驟清單、切換步驟、載入 ExcelJS
js/excel-io.js              所有 Excel 讀寫與排版（ExcelJS）
js/ui.js                    各步驟共用的畫面元件：上傳區、試算表預覽、下載，以及 WIP／月份行的共用設定
js/steps/step1-mps-filter.js    第 1 步　篩選 MPS 報表
js/steps/step2-mps-totals.js    第 2 步　匯總
js/steps/step3-p003-import.js   第 3 步　匯入 P003 標準工時
js/steps/step4-ct-lookup.js     第 4 步　帶入標準工時 CT.
js/steps/step5-demand-hours.js  第 5 步　計算需求工時
js/steps/step6-summary.js       第 6 步　需求工時匯總
js/steps/step7-manpower.js      第 7 步　匯入人力報表
js/steps/step8-report.js        第 8 步　總表
vendor/supabase.js          supabase-js 2.117.3（連 Supabase 用，必要）
vendor/                     ExcelJS 本機副本（可選）
icons/                      圖示
啟動小程式.command           在 Mac 上啟動
.gitignore                  不進 git 的檔案（.DS_Store、.env）
```

新增步驟時：在 `js/steps/` 加一個檔案、在 `js/app.js` 的 `STEP_FILES` 加上檔名（各步驟由 app.js 在登入後載入，不寫在 `index.html`）、
在 `sw.js` 的 `SHELL` 加上檔名並把版本號加 1；有記住的輸入就在 `STEP_SETTINGS` 寫上這一步用到的組名。
檔名開頭的數字就是步驟編號；在中間插入或拿掉步驟時，後面的檔案和組名要跟著改（例如拿掉 RMA 時，見下面「拿掉 RMA 之後的重新編號」）。

每次改版：把 `js/excel-io.js` 裡的 `VERSION`、`sw.js` 的 `VERSION` 和 `index.html` 裡每個 `?v=` 的數字一起加 1。
畫面右上角會顯示「版本 N」，可以用來確認瀏覽器跑的是不是最新的程式。

每一步做完會把結果放在 `IEApp.state`（第 1 步是 `mpsFiltered`，第 2 步是 `mpsTotals`，第 3 步是 `withP003`，第 4 步是 `withCT`，第 5 步是 `withDemand`，第 6 步是 `withSummary`，匯入人力報表是 `withManpower`，總表是 `withReport`），
裡面的 `build()` 會產生一份新的活頁簿，下一步拿來接著加工。

## 登入與共用設定（Supabase）

部門的人共用同一份「記住的設定」：誰最後存，大家打開看到的就是誰的。

- 打開時要先登入（email + 密碼）。帳號由管理者在 Supabase 後台建立（Authentication → Users → Add user），不開放自行註冊。
- 設定存在 Supabase 專案的 `ie_manpower.settings` 表：一筆是 `(app, key) → value`。
  - `app`：專案代號（`js/config.js` 的 `PROJECT_ID`）。8 個人力需求 PWA 共用這張表，用 `app` 分開。
  - `key`：「步驟.輸入項目」，例如 `step8.otHours` 是第 8 步總表的每日加班小時。組名一律用步驟編號（`step1`～`step8`），第 2、5 步共用的 WIP／月份行是 `step2-5`，第 1、2、6 步共用的 Status 行是 `step1-2-6`。
  - `value`：那個輸入項目的值（JSON）。6 個月、Off line 各種人數這種一起看的值放在同一筆（陣列）。
  - `updated_at`、`updated_by` 由資料庫的 trigger 自動填，程式不送。
  - 一個輸入項目一筆，所以兩個人同時改同一步的不同項目，不會互相蓋掉；每個項目都記得是誰改的。
- 誰改的、什麼時候改的由資料庫自動記錄（`updated_by`、`updated_at`），各步驟標題下面會顯示「共用設定：上次由 某某 於 10/19 16:05 修改」。
- 右上角顯示登入的人、同步狀態（已同步／同步中／離線）和「登出」。
- 讀設定一律讀這台電腦的本機備份（localStorage），所以離線也能用；改了設定先存本機，再在背景上傳（停止輸入 1 秒後）。離線時改的會等連上網路自動上傳。
- 打開時先把 Supabase 上最新的設定抓下來蓋掉本機備份，才載入各步驟（所以 `index.html` 裡沒有各步驟的 `<script>`，改由 `js/app.js` 的 `STEP_FILES` 載入）。
- 第一次登入時，這台電腦原本有、但 Supabase 上還沒有的設定會傳上去。
- 舊資料搬移（版本 7 以前只存在這台電腦的 localStorage）：第一次打開版本 8 時，舊的組名 `totals`、`columns`、`statusCol`、`step3`～`step6`、`manpower`、`report` 會在本機改成步驟編號（對照表在 `js/cloud.js` 的 `LEGACY_GROUPS`）；以前只記在第 1 步裡的 Status 行也會搬到 `step1-2-6`。新版不再存的 `step1.statusCol`、總表的 `effVer` 會拿掉。登入後，這台電腦有、但 Supabase 上還沒有的項目會自動上傳。
- 拿掉 RMA 之後的重新編號（版本 12）：RMA 原本是 `step7`、匯入人力報表原本是 `step8`、總表原本是 `step9`；現在匯入人力報表是 `step7`、總表是 `step8`。舊的內容意思不一樣，不搬，**第 7、8 步記住的設定要重新輸入一次**（人力報表各項目的列、總表的工作時間、效率、出勤率、加班、標題等）。
  - 每台電腦第一次打開版本 12 時，瀏覽器裡舊的 `step7`、`step8`、`step9` 會清掉（不會被當成新的設定傳上去），其他步驟的設定不受影響。
  - Supabase 上還留著的舊項目，同步時會略過不用：`js/cloud.js` 的 `GROUP_FIELDS` 寫明 `step7`、`step8` 各有哪些欄位，其他欄位（例如 RMA 的 `step7.items`、舊人力報表的 `step8.rows`）和整個 `step9` 都不理。**之後第 7、8 步新增記住的輸入項目時，要一起加到 `GROUP_FIELDS`**，不然不會同步。
  - 想把 Supabase 上的舊資料刪乾淨（可選，不刪也能用）：在 Studio 的 SQL Editor 執行（Mac、公司筆電的 Supabase 各做一次）。
    只會刪掉 `GROUP_FIELDS` 以外的舊項目，新版存的第 7、8 步設定不會被刪，什麼時候跑都可以：
    ```sql
    delete from ie_manpower.settings
    where app = 'ie-manpower-2F-automotive-assembly'
      and (key like 'step9.%'
        or (key like 'step7.%' and key not in ('step7.rows', 'step7.lastFirstCol'))
        or (key like 'step8.%' and key not in ('step8.title', 'step8.otHours', 'step8.otRate', 'step8.yoLimit',
              'step8.newModels', 'step8.reworkPeople', 'step8.offline', 'step8.workMinutes',
              'step8.effMass', 'step8.effTrial', 'step8.effRework', 'step8.attendance')));
    ```
    舊人力報表的 `step8.rows`、`step8.lastFirstCol` 不在總表的欄位裡，也會一起刪掉。

2F 組裝存進資料庫的輸入項目（`app` 都是 `ie-manpower-2F-automotive-assembly`）：

| key | 對應的輸入框 | value 範例 |
| --- | --- | --- |
| `step1.startRow`、`step1.selCol` | 第 1 步：資料起始列、「A 行等於」方式的選取條件在哪一行 | `5`、`"A"` |
| `step1.reports` | 第 1 步：各位置的條件（依 取像 → 車用 → 光機系統 排；同類別的第幾筆就是那一類的第幾份）、篩選方式和保留的行。`cat`：`image` 取像、`auto` 車用、`optics` 光機系統（以前沒有 `cat` 的舊資料，前三筆當作取像、車用、光機系統各一份）。`mode`：`a`（或沒有）= A 行等於，值在 `values`；`col` = 指定行，`fcol` 哪一行、`op` 是 `eq` 等於或 `neq` 不等於、值在 `fvalues`；`fz` = 模糊比對，`zCol` 哪一行、`zOp` 是 `eq` 選取或 `ne` 排除、`zValues` 各輸入框的值（陣列，輸入框有幾個就幾個值）；`blk` = 依區塊（TTL），`bCol` TTL 寫在哪一行（沒有就是 A）、`bBlocks` 勾的區塊名稱（陣列） | `[{"cat":"image","values":"1","cols":"B:K"},{"cat":"auto","mode":"col","fcol":"D","op":"neq","fvalues":"X","cols":"B:K"},{"cat":"optics","mode":"fz","zCol":"C","zOp":"eq","zValues":["Lens","Tray","Mech"],"cols":"B:K"}]` |
| `step1.statusPos` | 第 1 步：Status 是第幾個保留的行 | `"3"` |
| `step1.rma` | 第 1 步：挑出 RMA 的條件（所有報表共用）。`mode`：`fz` 模糊比對（值在 `zValues`，陣列）／`eq` 等於（值在 `values`，逗號分開）；`col`：新檔的第幾行（1 起算，空的 = 不挑）。取消勾選的列不記住 | `{"mode":"fz","col":"2","values":"","zValues":["RMA","",""]}` |
| `step1-2-6.statusCol` | 第 1、2、6 步共用：Status 在新檔的哪一行 | `"C"` |
| `step2.startRow` | 第 2 步：資料起始列（報表版面不一樣時） | `""` |
| `step2-5.wipCol`、`step2-5.monthCols` | 第 2、5 步共用：WIP 在哪行、各月在哪幾行 | `"D"`、`"E:J"` |
| `step4.keyCol`、`step4.refKeyCol`、`step4.refValueCol`、`step4.refDivCol`、`step4.refWcCol`、`step4.startRow` | 第 4 步：MPS 料號、P003 料號／Labor／Base Qty／Work Center 的行（`refWcCol` 是 `-` = 不帶入工作中心）、資料起始列 | `"A"`、`"B"`、`"M"`、`"K"`、`"N"`、`""` |
| `step5.startRow`、`step6.startRow` | 第 5、6 步：資料起始列 | `""` |
| `step7.rows`、`step7.lastFirstCol` | 第 7 步匯入人力報表：各項目在人力報表第幾列、上次的當月行 | `{"online":4,"offDirect":6}`、`4` |
| `step8.title` | 第 8 步總表：總表大標題 | `"製二課 (車用/光機)"` |
| `step8.workMinutes`、`step8.effMass`、`step8.effTrial`、`step8.effRework`、`step8.attendance` | 第 8 步總表：工作時間、綜合效率%（量產／試產／重工）、出勤率%，各 6 個月 | `["480","480",…]` |
| `step8.otHours`、`step8.otRate`、`step8.yoLimit` | 第 8 步總表：每日加班小時、加班出勤人數 %、YO 加班上限 | `"2"`、`"80"`、`"78"` |
| `step8.newModels`、`step8.reworkPeople` | 第 8 步總表：需求人力「新機種+重工」的新機種名稱和各月人數、重工各月人數 | `[{"name":"X100","qty":["2","2","3","3","3","3"]}]` |
| `step8.offline` | 第 8 步總表：需求人力「Off line」的項目和人數（項目可以增減；舊格式只有人數的陣列，讀的時候會套上預設的四個項目名稱） | `[{"name":"領班","qty":"1"},{"name":"品保","qty":"2"}]` |

取名原則：組名一律用步驟編號（總表也是 `step8`）；幾步共用的組寫成 `step2-5`、`step1-2-6`，看名字就知道是哪幾步共用（同一個軟件的幾步共用，不是各產線共用）。一個輸入項目一筆，6 個月、幾種人數這種一起看的值放在同一筆（陣列）。

- 不記住的欄位（上傳的檔案、第 5 步的月份表頭、第 1 步 RMA 取消勾選的列、第 8 步的平日工作天數）照舊每次輸入，不會上傳。
- 程式在 `js/cloud.js`，用 supabase-js 套件連線（`vendor/supabase.js`，版本 2.117.3，放在專案裡，離線也載得到）。雲端、本機自架、公司主機上的 Supabase 都用同一套程式，換資料庫只改 `js/config.js` 的 `SUPABASE_PUBLIC_URL`、`SUPABASE_PUBLISHABLE_KEY`。各步驟用 `IEApp.settings.read('step8')`／`IEApp.settings.write('step8', 物件)` 讀寫一整步的設定，`cloud.js` 負責拆成一個個輸入項目上傳、下載時再組回去。
- Supabase 的網址和公開鑰匙在 `js/config.js`。公開鑰匙本來就是公開用的；安全靠登入和資料表的權限規則（只有登入的人能讀寫，不能刪除）。`SUPABASE_SECRET_KEY`（舊式名稱：`SERVICE_ROLE_KEY`）不可以放進程式。
- 取名原則：`js/config.js` 的設定名稱照自架 Supabase 的 `.env` 取，兩邊找得到同一個名字。

  | `js/config.js` | 對應自架 `.env` | 說明 |
  | --- | --- | --- |
  | `SUPABASE_PUBLIC_URL` | `SUPABASE_PUBLIC_URL` | Supabase 網址 |
  | `SUPABASE_PUBLISHABLE_KEY` | `SUPABASE_PUBLISHABLE_KEY`（舊式名稱：`ANON_KEY`） | 公開鑰匙，`sb_publishable_` 開頭 |
  | `SUPABASE_SCHEMA` | `PGRST_DB_SCHEMAS` 清單裡的其中一個 | 這個程式用的 schema（`ie_manpower`） |

- 登入資訊由 supabase-js 存在這台電腦的 `ie-manpower-2F-automotive-assembly._auth`，會自己換新憑證。
- 有登入過的人打開時，如果 4 秒內連不上 Supabase，就先用這台電腦的設定，連上之後自動同步。

### 換一個 Supabase（本機自架、公司主機）

每個新的 Supabase 都要做一次：

1. 執行下面的 SQL（Studio 的 SQL Editor）。
2. 讓 API 看得到 `ie_manpower` schema：
   - 雲端：Settings → Data API → Exposed schemas 加上 `ie_manpower`。
   - Mac 上的 supabase CLI：`supabase/config.toml` 的 `[api]` 裡 `schemas` 加上 `"ie_manpower"`，再 `supabase stop` / `supabase start`。
   - Docker 自架：docker 資料夾 `.env` 的 `PGRST_DB_SCHEMAS` 加上 `,ie_manpower`，再 `docker compose up -d` 重啟。
3. 關閉自行註冊：CLI 是 `config.toml` 的 `[auth] enable_signup = false`；Docker 是 `.env` 的 `DISABLE_SIGNUP=true`。
4. 在 Studio 的 Authentication → Users 新增同事的帳號（email + 密碼，勾 Auto Confirm）。
5. 改 `js/config.js` 的 `SUPABASE_PUBLIC_URL`、`SUPABASE_PUBLISHABLE_KEY`（各環境要填什麼寫在檔案開頭）。

搬到公司主機時要注意：

- 別台電腦連的網址要寫主機的 IP 或名稱，不能寫 `127.0.0.1`／`localhost`（那是指使用者自己的電腦）。
- 用 https 網址打開的程式（例如 GitHub Pages）不能連 http 的 Supabase（瀏覽器會擋）。程式和 Supabase 要都是 http（例如程式也放在公司主機上），或都是 https。
- 離線快取（Service Worker）和「安裝成 App」只在 https 或 `localhost` 打開時才有；用 `http://主機IP` 打開時程式照常能用，但沒有離線功能。

建立資料庫的 SQL（雲端 Supabase 已經執行過；Exposed schemas 要加上 `ie_manpower`，見上面第 2 點）：

```sql
create schema if not exists ie_manpower;
grant usage on schema ie_manpower to authenticated;
create table if not exists ie_manpower.settings (
  app text not null, key text not null, value jsonb not null,
  updated_at timestamptz not null default now(), updated_by text,
  primary key (app, key)
);
alter table ie_manpower.settings enable row level security;
grant select, insert, update on ie_manpower.settings to authenticated;
create policy "settings_read"   on ie_manpower.settings for select to authenticated using (true);
create policy "settings_insert" on ie_manpower.settings for insert to authenticated with check (true);
create policy "settings_update" on ie_manpower.settings for update to authenticated using (true) with check (true);
create or replace function ie_manpower.stamp_settings() returns trigger language plpgsql security definer set search_path = '' as $$
begin new.updated_at := now(); new.updated_by := auth.jwt() ->> 'email'; return new; end $$;
create trigger settings_stamp before insert or update on ie_manpower.settings
for each row execute function ie_manpower.stamp_settings();
```

## 專案代號與複製成新專案

這個專案的代號是 `ie-manpower-2F-automotive-assembly`，寫在 `js/config.js`（網頁和 `sw.js` 都讀這個檔案）。

- 離線快取的名稱：`ie-manpower-2F-automotive-assembly-v8`（代號 + 版本號，版本號在 `sw.js` 的 `VERSION`）。新版啟用時，只刪「這個代號、但不是目前版本」的舊快取，別的專案的快取不碰。
- 記住的設定：Supabase 上是 `app` = 專案代號、`key` = `step1.startRow` 這種名稱；這台電腦的本機備份（localStorage）是 `ie-manpower-2F-automotive-assembly.step1` 這種名稱。各步驟一律用 `IEApp.settings.read('step1')`／`IEApp.settings.write('step1', 物件)` 讀寫，不直接碰 localStorage。
- 這個專案是從 3F 取像版（`ie-manpower-3F-image-lens`）複製出來的，沒有舊名稱要搬（`config.js` 的 LEGACY_* 都是 null），不會動到 3F 取像版的設定和快取。
- 啟動用的 port 是 8766（3F 取像版是 8765）。兩個專案開在同一個 port 會被瀏覽器當成同一個網站，離線快取的背景程式會互相蓋掉。

複製成新專案時要改的地方：

1. `js/config.js`：`PROJECT_ID` 改成新專案的代號（`ie-manpower-<產線>`）；`LEGACY_STORE_PREFIX`、`LEGACY_CACHE_PATTERN` 改成 `null`。`SUPABASE_*` 不用改（所有產線共用同一個 Supabase、同一個 schema，用 `PROJECT_ID` 分開）。
2. `manifest.webmanifest`：`name`、`short_name`（最好也換 `theme_color` 或圖示）。
3. `啟動小程式.command` 的 `PORT` 和 `.vscode/settings.json` 的 port 改成同一個號碼，各專案從 3001 開始往上加（3F 取像 3001、2F 組裝 3004…），本機測試時兩個專案才不會是同一個網址。
4. `sw.js` 的 `VERSION`、`js/excel-io.js` 的 `VERSION`、`index.html` 的 `?v=` 可以從 1 重新開始。
