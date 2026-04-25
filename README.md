**(日本語は下にあります)**

Prefer offline reading? [Download the PDF Manual here!](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper/blob/main/KENZEN_SeaArt_Helper_Manual.pdf)

**■ KENZEN SeaArt Helper Manual (NSFW Prompt Builder)**

**\* Summary**

**This is a professional prompt builder for SeaArt / Stable Diffusion, specialized in NSFW content and Japanese subculture styles.**

**1\. Introduction**

**This Excel workbook is powered by VBA macros.**

- **Co-Developed with: Google Gemini (The heavy lifter who did the actual coding).**
- **Open Source Ethos: Press Alt+F11 to open the VBA editor. Feel free to inspect, tweak, or overhaul the logic in the MainCode module.**
- **Copyright & License: I retain the original copyright, but you are free to modify and redistribute it. No permission required (though a shout-out makes the author happy).**

**2\. Design Philosophy**

**Developed to drastically min-max the efficiency of NSFW-focused prompt construction in SeaArt. This is a dedicated English prompt builder designed to bridge the gap between your imagination and the AI's output. It's packed with my personal passion and... specific "preferences." Avoid the pitfalls of "Engrish" and take the shortest path to your ideal generation. ……Don't worry, it's totally KENZEN (Wholesome)!**

**3\. Core Features (Safety First)**

- **Full-Width Guard: If Japanese (full-width) characters are detected, the tool stops the copy process and alerts you with a beep.**
- **Auto-Concatenation: Starting from the second copy, the tool automatically inserts a , (comma + space) between tags.**
- **Real-Time Debugging: Your active prompt is always displayed in cell C2.**
- **One-Click Reset: The Clear button wipes both the clipboard and the C2 cell instantly.**

**4\. Getting Started**

- **Unblock the File: Right-click the downloaded .xlsm file → Properties → Check "Unblock" at the bottom → Apply.**
- **Enable Macros: Open the file and click "Enable Content" on the yellow security warning bar.**

**5\. Search Panel & Controls**

- **Floating Window: Stays "Always on Top" for a seamless workflow.**
- **Smart Search:**
  - **Enter a keyword and click Search (or hit Enter).**
  - **The screen flashes for a split-second as the cursor teleports to the result.**
  - **Hit Enter repeatedly to loop through multiple matches.**
- **Refinement: \* Prompts accumulate in C2 and C4 (or C3 on the Sample sheet).**
  - **Cell C4: Use this for manual tweaks (adding custom tags not in the list). Select this cell and click Copy to finalize.**
- **Special Operations:**
  - **Copy without Comma: Connects tags with only a space. Perfect for multi-word concepts like oversized tank top.**
- **Undo & Purge:**
  - **Undo: Roll back your last action (1-step buffer).**
  - **Clear: Wipes clipboard and C2 (keeps your manual C4 tweaks).**
  - **All Clear: Hard reset. Wipes everything.**

**6\. The "Weighting" Engine**

**Toggle the Weight checkbox in the floating window, select a value (1.1 - 1.3), and click Wrap Block. This wraps everything up to the previous comma in parentheses-e.g., (oversized tank top:1.1). _Safety Feature:_ To prevent accidental double-weighting and reduce AI compute load, the checkbox auto-resets to OFF after one use.**

**7\. Workflow Flowchart**

**Select cells from left to right to build a complete narrative:**

**Character Count → Skin & Attributes → Hair length → Bangs → Tying → Hair Color → Body Type → Occupation → Outfit → Outfit State → Footwear & Legwea r→ Accessories → Location → Time & Surroundings → Position → Action & Movement → Means & Props → Body Parts → Interaction State → Expressions & Physical States → Misc Items → Camera Angle → Censorship Fixes**

**8\. Mastering "Copy without Comma"**

**The macro defaults to adding a comma. For phrases where a comma would break the logic-e.g., 1girl, walking in park with me,-do the following:**

- **Copy 1girl (Normal).**
- **Copy walking (Normal).**
- **Copy in (Without Comma).**
- **Copy park (Without Comma). ...and so on. Use Normal Copy only when you want the comma separator back.**

**9\. "Add to Fav" Button**

**This button allows you to transfer your favorite prompts, along with a description, to the "My Favorite" sheet. When you click the button, an input screen for the description will appear. Simply write a note that is easy for you to understand and click "OK" to register the entry. You can store up to 50 items. Please note that you cannot register a prompt if the description field is left blank.**

**10\. "Sample Prompts" Sheet**

- **The "Fetish Reveal" (Samples): I've included my favorite scenarios in the "Sample Prompt" tab. Use them as a base or a source of... inspiration.**

**Note: Manual weighting is required for custom samples.**

- **Negative Prompting: Don't choke the AI. Over-tagging negatives causes "Prompt Bleed" or glitches. Keep it minimal (e.g., just enough to banish "unwanted males").**

**11\. "My Favorite" Sheet**

**This sheet allows you to manage up to 50 of your favorite prompts.**

- **"Fav Copy" Button: Clicking this copies the prompt to your clipboard and simultaneously displays its description in the "Selected Prompt" cell. (After all, who can actually recall the details of a long, complex incantation just by glancing at it?)**
- **"Fav Clear" Button: This clears only the current clipboard content and the "Selected Prompt" cell. (Think about it-you're working hard to build a new prompt on the main sheet; it would be a disaster if that got wiped out too, right?)**
- **Search Function: Enter a keyword in the search box at the top right and click the "Fav Search" button to search through your favorites. The search is limited to the "Description" cells, a design choice made to prevent confusion even if the same words appear in both a prompt and its description.**

**12\. Exporting and Importing Prompts**

**Exporting Behavior**

- **Format: Data is exported in CSV format with UTF-8 encoding.**
- **Filename: The default name is automatically set as "MyFavorite_yyyymmdd_HHmm" (Date and Time).**
- **Overwriting Protection: If a file with the same name already exists in the destination folder, a suffix like "(1)" or "(2)" is automatically added. This prevents accidental deletion of existing backups.**

**Importing Behavior**

- **Import Mode Selection: When executed, you can choose to either "Append to existing data" or "Clear all and import as new."**
- **Automatic Data Allocation:**
  - **If a row contains two or more values: The 1st value is assigned to "Description" and the 2nd value to "Prompt."**
  - **If only one value is found: It is forcibly stored in "Description" regardless of its original column, ensuring it remains searchable.**
- **Sliding (Packing) Logic: Even if the CSV contains invalid formats or empty rows, only valid data is extracted and imported sequentially from the top without leaving any gaps.**
- **Japanese Language Support: Full support for prompts and descriptions containing Japanese characters (Double-byte characters).**
- **Error Logging: If any rows are skipped due to invalid characters (e.g., control characters), a detailed error log (ImportLog_HHmmss.txt) will be generated in the same folder as the CSV.**

**13．"Author's Notes (Tips and Rants) " Sheet**

**Includes some tips and the "struggle stories" behind this macro. Good for a break.**

**14\. "CONTACT" Sheet**

**As also mentioned in this README, this sheet contains information such as the author's contact details and blog address.**

**15\. Credits & Disclaimer**

- **Disclaimer: Generation results are at the mercy of the AI. The author takes no responsibility for any damages (or lack of "nut") resulting from this tool.**
- **Tested on: RIN Anim8Draw Illustrious - Anime Drawing Model (Ver.4.0A)**

[**https://www.seaart.ai/ja/models/detail/d158n1te878c73atvtdg**](https://www.seaart.ai/ja/models/detail/d158n1te878c73atvtdg)

- **Author: Tomohito Fujikawa (aka "Dst" or "Deeste" / Former Eroge(Visual Novels) Writer).**
- **Blog/Support:** [**https://dsblog.biz/**](https://dsblog.biz/)
  - **Feedback and requests are welcome via the blog's mail form.**
  - **Tips are greatly appreciated via PayPal:** [**https://paypal.me/dst0508**](https://paypal.me/dst0508)
- **Bonus: Craving some Lore? Check out my SeaArt page! I put my VN writing skills to work by posting original short stories alongside my generated art.**

[**https://www.seaart.ai/ja/new-user/4b23d22e331a382c4adc23a3df4e7077**](https://www.seaart.ai/ja/new-user/4b23d22e331a382c4adc23a3df4e7077)

**ENJOY! :)**

----------------------

**■KENZEN SeaArt Helper マニュアル**

オフラインのマニュアルがご希望ですか？ [PDF版マニュアルはこちらからダウンロード！](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper/blob/main/KENZEN_SeaArt_Helper_Manual.pdf)

**要約：SeaArtでの、NSFW絵の生成プロンプト構築に特化したツールです。**

**1\. はじめに**

**当Excelブックにはマクロを使用しております。**

- **コーディング支援： Google Gemini（実質的な実装担当という名の丸投げ）**
- **透明性の確保： Alt+F11 でVBAエディタを開き、\[標準モジュール\] 内の MainCode を参照・改変いただけます。**
- **著作権： 放棄しませんが、改変および再配布は自由です。報告も不要です（あると作者が喜びます）。**

**2\. 作成目的**

**SeaArtでの「NSFW絵に特化した」プロンプト作成を劇的に効率化するために開発しました。 AIへの意図を正確に伝えるための、英語プロンプト構築ツールです。「自分用に使いやすいものを！」という情熱と性癖を詰め込みました。「ダメ英語」の落とし穴を回避しつつ、最短ルートで理想の出力を目指します。……全き！　健全（KENZEN）ですよ！**

**3\. マクロの主要機能（安全設計）**

- **全角チェック： 日本語（全角文字）が含まれている場合、コピーを停止し警告音で知らせます。**
- **自動連結： 2回目以降のコピー時、自動的に ", "（カンマ＋半角スペース）を挿入して追記します。**
- **リアルタイム表示： 構築中のプロンプトはC2セルに常時表示されます。**
- **リセット機能： Clear ボタンでクリップボードとC2セルの内容を消去します。**

**4\. 使用準備**

- **ブロック解除： ダウンロードしたxlsmファイルを右クリック → プロパティ→ 「許可する」にチェックを入れて保存。**
- **マクロ有効化： ファイルを開き、上部の「コンテンツの有効化」をクリック。**

**5\. 検索パネルと操作方法**

- **高度な検索機能： 起動時に常に最前面に表示されます。**
  - **言葉を入力して Search（またはEnter）をクリック。**
  - **結果は一瞬点滅し、カーソルが移動します（全角検索なら右隣のプロンプトセルへ、半角ならその場に留まります）。**
  - **連続ヒットする場合、Enterを押すごとに次の結果へループ移動します。**
- **コピーと微調整：**
  - **コピーを繰り返すとC2およびC4（サンプルシートはC3）セルにプロンプトが蓄積されます。**
  - **C4セル： リストにないプロンプトを手入力するなどの調整後、このセルを選択して Copy を押せば反映されます。**
- **特殊コピー：**
  - **Copy without comma：カンマなし（半角スペースのみ）で連結します。oversized tank top 等、連結して一つの概念を指す場合に有効です。**
- **アンドゥとクリア：**
  - **Undo：1回だけ操作を戻せます。**
  - **Clear：クリップボードとC2を消去（C4の微調整内容は保持）。**
  - **All Clear：全てを完全に消去し、リセットします。**
- **重み付け機能**
  - **Wrap Block：強調したいプロンプトを、1.1～1.3まで重み付けできます。一つの単語はもちろん、例えば、先ほどの「oversized + tank top」といった、2単語以上の組み合わせも、「(oversized tank top:1.1)」のようにできます。**

**6\. プロンプトの構築フロー**

**薄緑色のセルを左から順に選んでいくだけで、一つの完成されたプロンプトになります。**

**キャラ数(Character Count) → 肌の色・属性(Skin & Attributes) → 髪の長さ(Hair length) → 前髪(Bangs) → 髪の結び目(Tying) → 髪の色(Hair Color) → 体型(Body Type) → 職業(Occupation) → 服装(Outfit) → 服の状態(Outfit State) → 足元周り(Footwear & Legwear) → アクセサリー類(Accessories) → 場所(Location) → 時間帯・周囲の状況(Time & Surroundings) → 体位(Position) → 行為・動作(Action & Movement) → 手段・道具(Means & Props) → 身体の部位(Body Parts) → 行為の状態(Interaction State) → 表情・生理現象(Expressions & Physical States) → その他アイテム(Misc Items) → アングル(Camera Angle) → 修正(Censorship Fixes)**

**7\. 「Copy without comma」の使い方**

**マクロは、「最初を除き、コピーされる単語の頭に、カンマと半角スペースを付ける」挙動をします。なので、例えば、「（接頭句としてのポジティブプロンプト）, "1 girl" "go to"" park" "with" "me"」の場合は、まず、「1 girl」で通常コピー、次に「go to」でも通常コピー、次の「park」でカンマなしコピー、その次の「with」でも、カンマなしコピー……と、「次にカンマを入れるべき所」（例の場合は「me」）まで、カンマなしコピーをしてください。次のプロンプトを通常コピーすれば、区切りに「, 」が付きます。**

**8\. 重み付け（ウェイティング）機能の使い方**

**フロートウィンドウ（メイン、検索共通）の、「Weight（ウェイト）」のチェックボックスをオンにして、プルダウンメニューから度合い（1.1～1.3）を選び、「Wrap Block（ブロックをくくる）」をクリックすると、「その前のカンマまでの単語全て」が、まとめてカッコでくくられて、重み付けされます。もちろん、1つのプロンプトにも有効です。誤操作防止と、AIへの負荷軽減のために、一度「Weight」ボタンをクリックすると、チェックボックスはオフになります。**

**9.** **Add to Favボタン**

**気に入ったプロンプトを、説明と共に、「My Favorite」シートに転送することができます。ボタンをクリックすると、説明書きの入力画面になるので、自分が分かりやすいような説明を書いて、「OK」をクリックすると、登録されます。最大50件までです。説明が空欄だと、登録できません。**

**10\. 「Sample Prompt」（サンプルプロンプト）シート**

- **サンプル： 「サンプルプロンプト」タブに作者お気に入りのシチュエーションを収録しています。性癖の開示！（電波）もちろんと言うべきか、それをコピーした上で、自分で調節して、オリジナルのプロンプトを作る事も出来ます。ただし、その場合の重み付けなどは、ご自身でお願いします。**
- **ネガティブプロンプト： 縛りすぎはAIのバグを誘発します。最小限に留めるか、いっそオフにするのがコツです（「野郎の出演」等、絶許な結果を防ぐ程度に）。**

**11．「My Favorite」シート**

**50件までの、お気に入りプロンプトの管理シートです。「Fav Copy」をクリックすると、「プロンプトがクリップボードにコピーされると共に、「説明書きが」「Selected Prompt」のセルに表示されます。（長々とした呪文だけ見て、詳細をすぐに思い出せる人も、そうはいないでしょう？）「Fav Clear」ボタンは、現在のクリップボードの内容と、「Selected Prompt」のセルのみを消去します。（いや、せっかくメインシートで新たなプロンプトを構築しているのに、それまで消えたら台無しじゃないですか？）右上の検索ボックスにキーワードを入力し、「Fav Search」ボタンをクリックすると、「My Favorite」内を検索できます。検索対象は「Description」セルのみですので、プロンプトと説明に同じ単語があっても、混乱しないような設計にしています。**

**プロンプトのエクスポートとインポート**

**「Export as CSV」及び、「Import from CSV」で、お気に入りリストのエクスポートとインポートができます。**

**・エクスポート（書き出し）の挙動**

**保存形式: 文字コード UTF-8 のCSV形式で出力されます。**

**ファイル名: デフォルトで「MyFavorite_yyyymmdd_HHmm（日付と時刻）」が設定されます。**

**上書き防止: 保存先に同名のファイルが存在する場合、自動的に「(1)」「(2)」といった枝番が付与されます。既存のバックアップを誤って消去することはありません。**

**・インポート（取り込み）の挙動**

**取り込みモードの選択: 実行時に「既存のデータに追記」するか、「全クリアして新しく取り込む」かを選択できます。**

**データの自動分配:1つの行に2つ以上の値がある場合：1つ目を「Description（説明）」、2つ目を「Prompt（プロンプト）」へ振り分けます。**

**値が1つしかない場合：データの位置に関わらず、検索の利便性を優先して強制的に「Description」へ格納します。**

**スライド（詰め）機能: CSV内に形式不良や空行があっても、有効なデータだけを抽出して上方向へ隙間なく詰めてインポートします。**

**日本語のサポート: 日本語（全角文字）を含むプロンプトや説明も完全にサポートしています。**

**エラーログ: 制御文字などの不純物により取り込めなかった行がある場合、CSVと同じフォルダに詳細なエラーログ（ImportLog\_時刻.txt）が生成されます。既存のリストに追加した結果が50件を超える場合、あふれたデータは、エラーログに記録されます。**

**「Fav All Clear」ボタンをクリックすると、現在のクリップボードの内容、お気に入りリスト、「Selected Prompt」のセル全てを消去します。**

**12．「作者覚え書き(ja) / Author's Notes (Tips and Rants)」シート**

**ちょっとしたTipsとか、このマクロを作るに当たっての苦労話とかを書いています。息抜きにどうぞ（？）**

**13．「CONTACT」シート**

**このREADMEにも書いていますが、作者の連絡先ブログなどを記載しています。**

**14\. 免責事項・連絡先**

- **生成結果はAI次第です。当ツールの使用による損害について、作者は一切の責任を負いません。**
- **検証モデル：RIN Anim8Draw Illustrious - Anime Drawing Model(Ver.4.0A)**

[**https://www.seaart.ai/ja/models/detail/d158n1te878c73atvtdg**](https://www.seaart.ai/ja/models/detail/d158n1te878c73atvtdg)

- **作者：不二川巴人（ふじかわ ともひと）（「でぇすて」とか、「不二川"でぇすて"巴人」名義で、エロゲーライターをやっていました）**
- **連絡先・ブログ：** [**https://dsblog.biz/**](https://dsblog.biz/)
  - **リクエストや感想はブログのメールフォームまで。投げ銭（PayPal）も歓迎です！**

[**https://paypal.me/dst0508**](https://paypal.me/dst0508)

**（おまけ）**

- 　**SeaArtの個人ページでは、生成したイラストを元に、書き下ろしショートショートを投稿したりしています。よろしければ、そちらもどうぞ。**

<https://www.seaart.ai/ja/user/4b23d22e331a382c4adc23a3df4e7077?u_code=XWACJSXI>

**健やかなる()AIライフを！**
