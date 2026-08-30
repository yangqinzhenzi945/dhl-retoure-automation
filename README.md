# DHL Retoure Automation

浠?Excel 璇诲彇閫€璐ф暟鎹紝鍦?DHL Gesch盲ftskundenportal 涓壒閲忓垱寤?`DHL Retoure Online` 鏍囩锛屽苟鎶婄粨鏋滈〉 `Sendungsnummer` 鍐欏洖 Excel 鐨?L 鍒椼€?
## 閲嶈鎻愮ず

- 鏈伐鍏蜂細鍒涘缓鐪熷疄 DHL 閫€璐у鎵橈紝璇峰厛浣跨敤 `--dry-run` 妫€鏌ョ涓€琛屻€?- 鐢ㄦ埛鍚嶃€佸瘑鐮佸拰鐧诲綍 Cookie 涓嶄細鍐欏叆椤圭洰鎴?Excel銆?- 鐧诲綍浼氳瘽淇濆瓨鍦ㄥ綋鍓嶇數鑴戠敤鎴风洰褰曠殑 `.dhl-retoure-automation/chrome-profile` 涓€?- 宸叉湁 `Sendungsnummer` 鐨勮浼氳嚜鍔ㄨ烦杩囷紝鏀寔涓柇鍚庣户缁€?- 姣忔垚鍔熷垱寤轰竴鍗曞氨绔嬪嵆淇濆瓨 Excel锛屽噺灏戜腑鏂鑷寸殑鏁版嵁涓㈠け銆?- DHL 椤甸潰缁撴瀯鍙樺寲鍚庯紝閫夋嫨鍣ㄥ彲鑳介渶瑕佹洿鏂般€?
## 鐜瑕佹眰

- Windows 10/11锛堝綋鍓嶇増鏈富瑕佸湪 Windows + Chrome 涓嬩娇鐢級
- [Node.js 20 鎴栨洿鏂扮増鏈琞(https://nodejs.org/)
- Google Chrome
- 鏈夋潈闄愪娇鐢?DHL Gesch盲ftskundenportal 鐨勮处鎴?
## Excel 琛ㄥご

绗竴寮犲伐浣滆〃绗?1 琛屽繀椤诲寘鍚互涓嬫爣棰橈細

| Excel 鏍囬 | DHL 瀛楁 |
|---|---|
| Sendungsreferenz | Sendungsreferenz |
| KUNDENREFERENZ / AUF LABEL ANZEIGEN | Kundenreferenz |
| VOR-UND NACHNAME | Vor- und Nachname |
| NAMENSZUSATZ 1 | Namenszusatz 1 |
| NAMENSZUSATZ 2 | Namenszusatz 2 |
| PLZ | PLZ锛堝洓浣嶉偖缂栬嚜鍔ㄨˉ 0锛?|
| ORT | Ort |
| STRASSE | Stra脽e |
| NR | Hausnummer |
| E-MAIL ADRESSE DES KUNDEN | Kunden-E-Mail |
| Retouren-Empf盲nger | DHL 璐︽埛涓殑鏀朵欢鏂瑰悕绉?|
| Sendungsnummer | 杈撳嚭鍒楋紱娌℃湁鏃惰嚜鍔ㄥ垱寤轰负 L 鍒?|

## 蹇€熷紑濮嬶紙Windows锛?
1. 涓嬭浇骞惰В鍘嬮」鐩€?2. 鍙屽嚮 `start.bat`銆傜涓€娆¤繍琛屼細鑷姩瀹夎渚濊禆銆?3. 灏?Excel 鏂囦欢鎷栬繘鍛戒护绐楀彛骞舵寜 Enter銆?4. Chrome 鎵撳紑鍚庤嚜琛岀櫥褰?DHL銆?5. 纭寰呭鐞嗘暟閲忥紝杈撳叆 `CREATE` 鎵嶄細寮€濮嬬湡瀹炴壒閲忓垱寤恒€?6. 杈撳嚭鏂囦欢榛樿淇濆瓨鍦ㄦ簮 Excel 鍚岀洰褰曪紝鏂囦欢鍚嶅鍔?`_宸插～鍐欒繍鍗曞彿`銆?
## 鎺ㄨ崘锛氬厛璇曡繍琛?
```powershell
npm install
npm start -- --input "C:\data\DHL閫€璐ф爣绛?xlsx" --dry-run
```

璇曡繍琛屽彧濉啓绗竴鏉″緟澶勭悊璁板綍锛屼笉鐐瑰嚮 `Retoure beauftragen`銆?
纭鏃犺鍚庤繍琛岋細

```powershell
npm start -- --input "C:\data\DHL閫€璐ф爣绛?xlsx"
```

鍙€夊弬鏁帮細

- `--output "璺緞.xlsx"`锛氭寚瀹氳緭鍑烘枃浠躲€?- `--overwrite`锛氱洿鎺ヨ鐩栬緭鍏ユ枃浠讹紙榛樿涓嶄細瑕嗙洊锛夈€?- `--dry-run`锛氬彧濉涓€鏉″苟鍋滃湪鎻愪氦鍓嶃€?- `--headless`锛氭棤鐣岄潰妯″紡锛涢娆＄櫥褰曚笉寤鸿浣跨敤銆?
## 闅愮涓庡畨鍏?
璇峰嬁鎶婄湡瀹炲鎴?Excel銆佹祻瑙堝櫒鐧诲綍鐩綍銆佹棩蹇楁垨瀵嗙爜鎻愪氦鍒?GitHub銆傞」鐩殑 `.gitignore` 宸插拷鐣?Excel銆佽緭鍑虹洰褰曞拰鏈湴娴忚鍣ㄤ細璇濄€?
## 鍏嶈矗澹版槑

杩欐槸闈炲畼鏂瑰伐鍏凤紝涓?DHL 鏃犻毝灞炲叧绯汇€備娇鐢ㄨ€呭簲纭繚鑷姩鍖栨搷浣滅鍚堝叾 DHL 鍚堝悓銆佸唴閮ㄦ祦绋嬪拰閫傜敤娉曞緥銆傚缓璁厛鐢ㄥ皯閲忔暟鎹祴璇曪紝骞朵汉宸ユ娊鏌ョ粨鏋溿€?
