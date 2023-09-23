# Foreclosure
Foreclosure Property Data Processing Software

## 軟體及函式庫
-  Python 3.10.9
-  openpyxl

## 用法
### **第一次使用步驟：（dist資料夾內無trace.xlxs、not_trace.xlxs）**
1. 下載法拍屋資料，將其存到dist資料夾內（若下載的檔案不是.xlxs，則更改為.xlxs），並更改名稱為：target.xlxs。
2. 執行主要程式(xlxs_vN.exe)(註：N為版本號，如：xlxs_v5.exe)
3. 執行完畢後，dist資料夾內，會產生trace.xlxs及not_trace.xlxs兩個檔案。
4. 再執行一次主要程式(xlxs_vN.exe)，此時會將target中新增兩個分頁，分別為「追蹤檔案」及「不追蹤檔案」。
5. 環境架設完畢，即可開始篩選。
6. 篩選完畢後，執行save程式（save_vN.exe），將篩選後的分頁內容，存到dist內的trace.xlxs、not_trace.xlxs。

第一次使用，防毒軟體可能會阻擋程式執行（因為需要撰寫檔案），要設定為允許，才能順利執行。


### **非第一次使用：（dist資料夾內已有trace.xlxs、not_trace.xlxs檔案）**
1. 下載法拍屋資料，將其存到dist資料夾內（若下載的檔案不是.xlxs，則更改為.xlxs），並更改名稱為：target.xlxs。
2.  執行主要程式(xlxs_vN.exe)，會根據dist資料夾內的trace.xlxs、not_trace.xlxs兩個檔案進行篩選。
3.  程式執行完後，target中，剩下的會是未曾瀏覽過的資料。
4.  開始篩選。
5.  篩選完畢後，執行save，將「追蹤檔案」及「不追蹤檔案」分頁內容，存回dist資料夾內的trace.xlxs、not_trace.xlxs。

