# Foreclosure
Foreclosure Property Data Processing Software

## 軟體及函式庫
-  Python 3.10.9
-  openpyxl
  
## 功能說明
每次下載法拍屋資料，都會有上千筆資料，其中不少資訊和之前下載的內容重覆，可能只有小部分修改(如: 拍賣次數等)，這些重複資訊讀起來會浪費不少時間，實際上新的法拍屋資料可能只有500筆，每次讀資料時，若只需要讀這些新資訊，就能省下不少時間。因此這份程式能夠幫助使用者篩選重複的資訊，並即時更新為最新資訊，讓使用者每次都只需注意新資料及有興趣的資料。

## 檔案說明
trace.xlxs: 此檔案儲存有興趣的資料，每次執行程式時，都會更新這份檔案該筆資料，使其成為最新資訊，日後只須注意這份檔案即可。
not_trace.xlxs: 此檔案儲存不感興趣之資料，美次執行程式時，以這份檔案內容進行篩選，刪除這些資料。

## 步驟說明 
1. 到 https://aomp109.judicial.gov.tw/judbp/wkw/WHD1A02.htm 下載法拍屋資料
2. 下載下來的檔案，應為xls檔案，需先轉換為xlsx，並改名為target
3. (**無trace及not_trace才需要此步驟!!**)
   
   執行xlxs_vN.exe，得到trace.xlxs及not_trace.xlxs，分別為希望追蹤的檔案，及希望刪除的檔案
-  trace.xlxs 內部的檔案會根據新下載的法拍屋資料進行更新(底價、拍賣次數、拍賣日期等都會更新為最新資訊)，並同步刪除target中該筆資料，確保每次資料都是未曾瀏覽過的
-  程式會根據 not_trace.xlxs 內的資料，將target中的資料刪除，避免該筆資料重複出現
4. 執行xlxs_vN.exe，以在target中產生兩分頁，分別為「追蹤檔案」及「不追蹤檔案」，以利篩選作業(若已有此兩分頁則可略過此步驟)。
5. 進行人工篩選，將每筆資料篩選至兩分頁。
6. 篩選完畢後，執行 save_vN.exe 將target中兩分頁的內容儲存至trace.xlxs及not_trace.xlxs中


![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/ded1e5e4-d74e-49ed-a4d3-a2b872a35d8f) 需要具備這些檔案，才能進一步執行
![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/8755cc60-ae25-466d-af9d-375d5fd4f1bf) 第一次執行程式，產生trace及not_trace
![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/f3b48f3c-a1fa-46bc-835b-2ad424bd5508) 產生的結果

![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/43ac8a9c-77a2-49e6-81fc-91e085ffb9b0) (上圖)再執行一次程式，產生「追蹤」及「不追蹤」分頁

![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/834d21c3-330f-4b42-bcc9-262b4df33be5)
 (上圖)複製資料至「追蹤」
![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/0fd58c4f-a684-4b69-ba95-b660f3925637)
 (上圖)複製資料至「不追蹤」
 
![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/64567090-5a4f-4112-9223-11c4c7081f9c) 儲存「追蹤」及「不追蹤」至trace及not_trace

![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/da8b5cb6-da24-4992-a19b-e05699dcfb89) (上圖)更改倒數第二筆資料，驗證後續的執行(改為第999拍)

![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/fa97daeb-90f2-4820-be51-4b5c7c8d0a3b) 執行程式，進行篩選，刪除並更新重複的資料
![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/df9b63d3-9400-4c4f-94a3-ffac565e36ed) (上圖)重複資訊皆刪除

![image](https://github.com/Chung-Wu/Foreclosure/assets/35622830/52473e42-3be1-42ac-af4a-147695a9a881) 更新資料為最新資訊(第999拍)





