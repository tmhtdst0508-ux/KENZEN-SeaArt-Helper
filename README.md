> # Welcome to the KENZEN SeaArt Helper! 🤝
> 
> **Greetings, my fellow People of Culture around the world!**
> 
> I am Tomohito Fujikawa (aka "Deesute"), a Japanese scenario writer. For about 15 years, I have crafted stories and situations for over 60 PC Visual Novels. In my pursuit of the ultimate "KENZEN" (Cultured/NSFW) AI art, I realized we needed a better way to manage our deepest desires and complex prompts. 
> 
> So, I built this tool for us. 
> 
> **KENZEN SeaArt Helper** is your ultimate cockpit for SeaArt and Stable Diffusion. 
> 
> 🚧 **[COMING SOON: v2.1.0 Major Update!]** 🚧
> I am currently doing final checks on version 2.1.0, which will introduce a highly requested **LoRA Management System**! Very soon, you will be able to set custom aliases, save recommended weights, stack multiple LoRAs, and save your favorite combinations as presets. Stay tuned!
> 
> Although I am primarily a writer, I poured my soul into coding this environment using AI, just so we can all break through our creative limits. 
> 
> Please use this tool to your heart's content, explore your fetishes, and create absolute masterpieces! 🎨
> 
> **If this tool helps you on your glorious journey, dropping a ⭐ Star on this repository would mean the world to me (and you'll be notified of the v2.1.0 release)!** Let's spread the culture together!

**(日本語は下にあります)**

**KENZEN SeaArt Helper Manual (v2.0.0)**

Prefer offline reading? [Download the PDF Manual here!](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper/blob/main/KENZEN_SeaArt_Helper_Manual_v2.0.0.pdf)

**TL;DR: A specialized prompt builder tailored for generating NSFW content in SeaArt (and Stable Diffusion).**

**1\. Introduction This Excel workbook is powered by VBA macros.**

- **Coding Assistance: Google Gemini (My unpaid AI intern who did the actual heavy lifting).**
- **Transparency: Press Alt+F11 to open the VBA editor and inspect/modify the MainCode module. _(Rest assured, there are no malicious backdoors. However, it does contain pure, artisanal spaghetti code cooked up exactly to my Google Gemini overlord's specifications.)_**
- **Copyright: I retain the original copyright, but you are free to modify and redistribute it. No permission is required (though a shout-out makes the author happy!).**

**2\. Design Philosophy Developed to drastically min-max the efficiency of NSFW-focused prompt construction in SeaArt. This is a dedicated English prompt builder designed to accurately bridge the gap between your intent and the AI's output. It’s packed with my personal passion and... specific "preferences." Avoid the pitfalls of "Engrish" and take the shortest path to your ideal generation. ...And let me repeat, it’s totally KENZEN (Wholesome(TM))!**

**3\. Core Features (Safety First)**

- **3-1. Anti-Double-Byte Alert: If Japanese (full-width) characters are detected, the tool stops the copy process and alerts you with a beep.**
- **3-2. Auto-Concatenation: Starting from the second copy, the tool automatically inserts , (comma + space) to append tags.**
- **3-3. Real-Time View: Your active prompt under construction is always visible in the "Current Prompt" field.**
- **3-4. Reset Function: The Clear button wipes both the clipboard and the screen instantly.**

**4\. Setup**

- **4-1. Unblock the File: Right-click the downloaded .xlsm file → Properties → Check "Unblock" → Apply/Save.**
- **4-2. Enable Macros: Open the file and click "Enable Content" on the security warning bar at the top.**

**5\. Search Panel & Controls Once the macro is active, the Main Window opens. Even if you accidentally close it with the top-right "X", you can reopen it anytime via the Open Window button(or the hotkey Ctrl + Shift + O). The window consists of two tabs (panels): Cockpit and Favorites.**

**5-1. The "Cockpit" Tab (Construction Zone) This tab contains the menus for building your prompt.**

- **5-1-1. Current Prompt: Your live workspace. Every copy action appends here. Since it's a text field, manual tweaks are also possible.**
- **5-1-2. Search Box (Enter Keyword): Type a keyword and hit Enter or click Search to scan the main sheet. When found, the cell glows yellow and the cursor jumps to the prompt cell on its right. (If searching in English, the cursor stays in place.) Hit Enter repeatedly to loop through multiple matches.**
- **5-1-3. "Set Positive Prompts" Button:This button automatically injects your quality-boosting prefix prompts (like masterpiece, best quality...) located in the "PositivePrompts" cell (A3) on the main sheet directly into your text field. You can also trigger this via the hotkey Ctrl + Shift + P. Sure, you _could_ manually select the cell and hit the "Copy" button mentioned later, but let's be real?that's a hassle, right? Nobody firing up an AI generator is looking for potato-quality art. We're here for masterpieces. If your preferred model/checkpoint requires a different set of quality tags, feel free to tweak the contents of cell A3. _Warning:_ Do NOT change the defined Name of the cell ("PositivePrompts"), or you'll break the macro's targeting system!**
- **5-1-4. Copy Button: Click this on a prompt cell to send it to the clipboard. As you repeat this, prompts accumulate in the "Current Prompt" field, separated by , (comma + space).**
- **5-1-5. Copy without comma Button: Connects tags with only a space. Perfect for compound concepts like oversized tank top. The macro defaults to adding a comma/space to the beginning of any copied word (except the first). For example: (positive prompt prefix), "1girl" "go to" "park" "with" "me". You would use Normal Copy for 1girl and go to, then use _Copy without comma_ for park and with... up until the point where a comma is needed again (me in this example). Using Normal Copy on the next prompt will resume adding the , separator. Note: Clicking this button for your very first prompt will not insert a leading space.**
- **5-1-6. Undo Button: Roll back your last action. Stores up to 50 steps. Your clipboard contents will also revert simultaneously.**
- **5-1-7. Clear Button: Wipes the clipboard and the "Current Prompt" field. (Undo is possible).**
- **5-1-8. All Clear Button: Total wipeout. Clears the clipboard, the "Current Prompt" field, and purges the entire history. (Undo is NOT possible). _(The "Nuke" / Tactical Obliteration Button)_**
- **5-1-9. Weighting Engine (Weight Checkbox & Wrap Block Button): Emphasize a prompt by weighting it from 1.1 to 1.3. Works for single words or multi-word blocks (e.g., (oversized tank top:1.1)). Clicking Wrap Block encloses everything up to the previous comma in parentheses and applies the weight. To prevent accidental double-weighting and reduce AI load, the checkbox auto-resets to OFF after one use. If you check the box again without adding a new prompt and click Wrap Block, the weighting will be removed.**
- **5-1-10. Done! Button: Once you're fully satisfied with your manual tweaks, click Done! to copy the final text field content. Since the clipboard already accumulates prompts during the building phase, this button remains inactive unless you manually edit the field.**
- **5-1-11. Send to Fav Button: Transfers your masterpiece prompt directly to the text field in the "Favorites" tab.**

**5-2. The "My Favorite" Tab (The Fetish Vault) The management menu for the precious prompts you've painstakingly forged. Note: None of the buttons in this tab will work unless the "My Favorite" sheet is actively displayed.**

- **5-2-1. Search Fav Field & Button: Enter a keyword and hit Enter/click Search Fav to scan the "My Favorite" sheet. It only searches within the "Description" field to prevent confusion. Matches glow and cycle just like the main search.**
- **5-2-2. Pull From Cockpit Button: Instantly grabs the prompt currently sitting in the Cockpit tab's text field.**
- **5-2-3. Send to Cockpit Button: The reverse of the above. Sends the contents of the Favorite Prompt text field to the Cockpit. Perfect for pulling up a base favorite and adding new elements from the main sheet to further refine it.**
- **5-2-4. Fav Clear Button: Clears all fields in the "Favorites" tab.**
- **5-2-5. Fav Copy Button: Copies your saved favorite from the sheet. After all, you want to cast your favorite spells over and over, right? The copied incantation (prompt) will be displayed in the text field along with its description. (Because let's face it, no one can remember what a massive wall of text actually does just by looking at it!)**
- **5-2-6. Tweaked! Button: After you use "Fav Copy" and manually adjust the prompt in the text field, this button lights up. Clicking it copies your newly refined version.**
- **5-2-7. Add to Fav! Button: Registers the prompt currently in the field to the "My Favorite" sheet. You cannot save it if the "Description" is blank. Maximum capacity is 50 slots.**
- **5-2-8. Replace Fav Button (Keep your library clean): Overwrites the currently focused cell on the "My Favorite" sheet with the contents of the Favorite tab's text field (if the cell is blank, it just inputs it). Trying to click this on a non-prompt cell will throw an error. Use case: "I tweaked my favorite prompt and it's finally perfect! But I don't want near-identical clones cluttering my list."**
- **5-2-9. Export Fav Button (Legacy of Fetishes: Preservation Protocol): Export your favorite list as a CSV file. We've all had that moment: "This generated image is amazing! Wait... what prompt did I use?!" Prompts are creative assets; this is your insurance policy to preserve them.**
    - **5-2-9-1. Exporting Behavior: Saves as a UTF-8 encoded CSV. Default filename is MyFavorite_yyyymmdd_HHmm. If a file with the same name exists, it automatically adds a suffix like (1) or (2) to prevent accidental overwrites of existing backups.**
- **5-2-10. Import Fav Button (Splicing Creative DNA): Import a favorite list from a CSV. Perfect for when you want to swap legendary prompts with your community.**
    - **5-2-10-1. Importing Behavior: You can choose to "Append to existing data" or "Clear all and import as new".**
        - **Automatic Allocation: If a row has 2+ values, the 1st becomes the "Description" and the 2nd becomes the "Prompt". If only 1 value exists, it is forcibly assigned to "Description" for searchability, regardless of its original column.**
        - **Packing Logic: Skips empty rows or bad formatting, extracting only valid data and stacking it neatly from the top.**
        - **Japanese Support: Fully supports Japanese descriptions.**
        - **Error Logging: Generates an ImportLog_HHmmss.txt if any rows are skipped due to control characters/impurities, or if the imported data exceeds the 50-slot limit.**
        - **_Note 1:_ The imported CSV must be UTF-8 encoded, or it will throw an error.**
        - **_Note 2:_ By design, it skips headers and starts importing from Row 4. Anything above that is ignored.**
- **5-2-11. Fav All Clear Button: Purges ALL data in the "My Favorite" sheet. This is irreversible, so use with extreme caution.**

**6\. The "Golden Flow" Workflow Simply select the light-green cells from left to right to construct a complete prompt: Character Count → Skin & Attributes → Body Type → Hair length → Bangs → Tying → Hair Color → Occupation → Underwear → Outfit → Outfit State → Headwear → Footwear & Legwear → Accessories → Location → Time & Surroundings → Position → Action & Movement → Means & Props → Body Parts → Interaction State → Expressions & Physical States → Misc Items → Camera Angle → Censorship Fixes. The hyperlinks at the top of the main sheet let you warp to specific categories. When clicked, the target cell automatically snaps to the far-left edge of the window. "Back to Legend" links are also provided at the end of each section.**

**7\. "Sample Prompts" Sheet Contains my personal favorite scenarios. Fetish disclosure alert! Copying a sample sends it straight to the "Cockpit" text box, allowing you to tweak it into your own original prompt. (Please handle any manual weighting yourself). _Tips: Negative Prompts_ Over-tightening chokes the AI and causes glitches. Keep it minimal or turn it off entirely (just enough to banish "unwanted males", for example).**

**8\. "Author's Notes (Tips and Rants)" Sheet Includes prompt tips and the struggle stories behind building this macro. Good for a break.**

**9\. "CONTACT" Sheet Contains the author's contact info and blog address, as listed in this README.**

**10\. Disclaimer & Contact Generation results are entirely at the mercy of the AI. The author takes no responsibility for any damages resulting from the use of this tool.**

- **AI results are unpredictable. I am not responsible for what you generate.**
- **Tested Model:** [**RIN Anim8Draw Illustrious - Anime Drawing Model (v4.0A)**](https://www.seaart.ai/ja/models/detail/d158n1te878c73atvtdg)
- **Tested LoRA：[Detailed anime style - Illustrious V2.0](https://www.seaart.ai/ja/models/detail/33ff0599ab0e8f9b66d7bee70551df23)**
- **Contact/Blog:** [**dsblog.biz**](https://dsblog.biz/)
- **Support the Dev: Bug reports are great, but** [**PayPal tips**](https://paypal.me/dst0508) **keep the lights on!**
- **While bug reports are absolutely welcome, what I _really_ crave are your missing tag requests! Hit me with feedback like, "Hey, you forgot this location!" or "Where the hell is this specific outfit?!" Sure, you have the freedom to mod the code and add them yourself, but please share them with me?because _I too wish to gaze upon uncharted scenarios!_ Your collective wisdom is the fuel that makes this macro even more totally KENZEN (Wholesome(TM))!**
- **Author: Tomohito Fujikawa (aka “Dst” or "Deeste" / Former Eroge (Visual Novels) Writer).**
- **Bonus: Craving some Lore? Check out** [**my SeaArt page!**](https://www.seaart.ai/ja/new-user/4b23d22e331a382c4adc23a3df4e7077) **I put my VN writing skills to work by posting original short stories alongside my generated art.**

**ENJOY your KENZEN AI Life! :)**

――――――――――――――――――――――――――――――――--

**■KENZEN SeaArt Helper マニュアル（v2.0.0）**

オフラインマニュアルは、[こちらをご覧下さい](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper/blob/main/KENZEN_SeaArt_Helper_Manual_v2.0.0.pdf)

**要約：SeaArt（＆Stable Diffusion）での、NSFW絵の生成プロンプト構築に特化したツールです。**

## **1.はじめに**

**当Excelブックにはマクロを使用しております。**

**コーディング支援： Google Gemini（実質的な実装担当という名の丸投げ）**

**透明性の確保： Alt+F11 でVBAエディタを開き、\[標準モジュール\] 内の MainCode を参照・改変いただけます。（※悪意あるバックドアは仕込んでいませんが、GoogleGeminiの言いなりで煮込まれたスパゲッティコードが内包されています）**

**著作権： 放棄しませんが、改変および再配布は自由です。報告も不要です（あると作者が喜びます）。**

## **2．作成目的**

**SeaArtでの「NSFW絵に特化した」プロンプト作成を劇的に効率化するために開発しました。 AIへの意図を正確に伝えるための、英語プロンプト構築ツールです。「自分用に使いやすいものを！」という情熱と性癖を詰め込みました。「ダメ英語」の落とし穴を回避しつつ、最短ルートで理想の出力を目指します。……全き！　健全（KENZEN）ですよ！**

## **3.マクロの主要機能（安全設計）**

### **3-1全角チェック**

**日本語（全角文字）が含まれている場合、コピーを停止し警告音で知らせます。**

### **3-2.自動連結**

**2回目以降のコピー時、自動的に “, “（カンマ＋半角スペース）を挿入して追記します。**

### **3-3.リアルタイム表示**

**構築中のプロンプトは「Current Prompt」フィールドに常時表示されます。**

### **3-4.リセット機能**

**Clear ボタンでクリップボードと画面表示内容を消去します。**

## **4．使用準備**

### **4-1ブロック解除**

**ダウンロードしたxlsmファイルを右クリック → プロパティ→ 「許可する」にチェックを入れて保存。**

### **4-2マクロ有効化**

**ファイルを開き、上部の「コンテンツの有効化」をクリック。**

## **5．検索パネルと操作方法**

**起動する（マクロを有効化する）と、メインウィンドウが開きます。**

**右上の×で閉じてしまっても、Open Main Window、あるいは、ショートカットキー Ctrl + Shift + O から、いつでも開けます。**

**ウィンドウは、「Cockpit」と「Favorite」の2つのタブ（パネル）で構成されています。**

### **5-1「Cockpit」パネル**

**プロンプトを構築していくためのメニューがあるタブです。**

#### **5-1-1.「Current Prompt」**

**ここに、現在のクリップボードの内容が表示されます。プロンプトをコピーするたびに、追記されていきます。テキストフィールドなので、手動での微調整も可能です。**

#### **5-1-2．検索ボックス（Enter Keyword）**

**フィールドにキーワードを入力し、エンターキー、ないしは「Search」ボタンをクリックすると、メインシート内を検索できます。ヒットすると、セルが黄色く光り、右隣のプロンプトのセルに移動します。英語で検索した場合は、その場にとどまり、移動しません。連続ヒットする場合、Enterを押すごとに次の結果へループ移動します。**

#### **5-1-3.「Set Positive Prompts」ボタン**

**メインシート上の「PositivePrompts」（A3）セルにある、「masterpiece…」から始まる接頭句としてのポジティブプロンプトを、テキストフィールドに自動入力します。この機能は、ショートカットキー Ctrl + Shift + Pでも実行できます。もちろん、シート上のセルを選択して、後述する「Copy」ボタンを押してもいいのですが、ぶっちゃけ面倒くさいじゃないですか？　AIに絵を描かせようって人間が、雑な絵を見たいはずがないですし。使用するモデルによって、ポジティブプロンプトの内容が違う場合は、各々適宜修正してください。ただし、「PositivePrompts」のセル名を変えると、機能しなくなります。**

#### **5-1-4.「Copy」ボタン**

**コピーしたいプロンプトのセル上でクリックすると、クリップボードに転送されます。コピーを繰り返すと「Current Prompt」フィールドに、「, 」（カンマと半角スペース）を付けて、プロンプトが蓄積されます。**

#### **5-1-5.「Copy without comma」ボタン**

**カンマなし（半角スペースのみ）で連結します。oversized tank top 等、連結して一つの概念を指す場合に有効です。マクロは、「最初を除き、コピーされる単語の頭に、カンマと半角スペースを付ける」挙動をします。なので、例えば、「（接頭句としてのポジティブプロンプト）, ”1 girl” ”go to”“ park” ”with” “me”」の場合は、まず、「1 girl」で通常コピー、次に「go to」でも通常コピー、次の「park」でカンマなしコピー、その次の「with」でも、カンマなしコピー……と、「次にカンマを入れるべき所」（例の場合は「me」）まで、カンマなしコピーをしてください。次のプロンプトを通常コピーすれば、区切りに「, 」が付きます。なお、一番最初のプロンプトを、このボタンでクリックしても、先頭にスペースは入りません。**

#### **5-1-6.「Undo」ボタン**

**操作を戻せます。50回まで可能です。クリップボードの内容も、元に戻ります。**

#### **5-1-7.「Clear」ボタン**

**クリップボードと「Current Prompt」フィールドを消去します。アンドゥは可能です。**

#### **5-1-8．「All Clear」ボタン**

**クリップボード、「Current Prompt」フィールド、及び全ての履歴完全に消去し、リセットします。アンドゥはできません。（通称：Nuke / 物理的更地化ボタン）**

#### **5-1-9．重み付け機能（「Weight」チェックボックス＆「Wrap Block」ボタン）**

**強調したいプロンプトを、1.1～1.3まで重み付けできます。一つの単語はもちろん、例えば、先ほどの「oversized + tank top」といった、2単語以上の組み合わせも、「(oversized tank top:1.1)」のようにできます。「Wrap Block」を押すと、その前のカンマまでのブロックを「()」でくくって、重み付けします。もちろん、1つのプロンプトにも有効です。誤操作防止と、AIへの負荷軽減のために、一度「Weight」ボタンをクリックすると、チェックボックスはオフになります。また、プロンプトを追加しないで、もう一度チェックをオンにして「Wrap Block」をクリックすると、重み付けが解除されます。**

#### **5-1-10.「Done!」ボタン**

**気が済むまで（？）手動での調整が終わったら、「Done!」をクリックすれば、現在のテキストフィードの内容がコピーできます。クリップボード内には、既にプロンプトが蓄積されていますから、編集をしない限り、このボタンはクリックできません。**

#### **5-1-11. Send to Favボタン**

**気に入ったプロンプトを、「Favorite」タブウィンドウのテキストフィールドに転送します。**

### **5-2.「My Favorite」タブ（性癖の金庫室）**

**苦心の末に（？）作り上げた、大切なプロンプトを管理するためのメニュー類があるタブです。なお、このタブのボタンは全て、「My Favorite」タブがアクティブになっていないと、動作しません。**

#### **5-2-1.「Search Fav」フィールド＆「Search Fav」ボタン**

**フィールドにキーワードを入力して、エンターか「Search Fav」ボタンのクリックで、「My Favorite」シート内を検索できます。検索されるのは「Description」のフィールドのみで、メインシートでのそれ同様、ヒットしたセルが光り、複数ヒットした場合は、ループします。**

#### **5-2-2.「Pull From Cockpit」ボタン**

**「Cockpit」タブのテキストフィールドに入力されているプロンプトを転写します。**

#### **5-2-3.「Send to Cockpit」ボタン**

#### **上記とは逆に、Favorite Promptのテキストフィールドの内容を、Cockpitのフィールドへ送ります。お気に入りを、メインシート内のプロンプトを加えることで、更にブラッシュアップするときに。**

#### **5-2-4.「Fav Clear」ボタン**

**「Favorite」タブのフィールドを、全てクリアします。**

#### **5-2-5.「Fav Copy」ボタン**

**「My Favorite」シート上のお気に入りをコピーします。やっぱりね、気に入った呪文は、何度でも使いたいですよね。「コピーされた呪文（プロンプト）は、説明と共にテキストフィールドに表示されます。（説明なしで呪文だけ見て、すぐに思い出せる人も、まずいないでしょう？）**

#### **5-2-6.「Tweaked!」ボタン**

**コピーしたお気に入りプロンプトを、テキストフィールド内にて、さらに手動で調整した後、クリック可能になり、クリックすると、テキストフィールドの内容がコピーされます。**

#### **5-2-7.「Add to Fav!」ボタン**

**フィールド内のプロンプトを、「My Favorite」シートに登録します。「Description」が空欄だと、登録できません。最大50件まで登録できます。**

#### **5-2-8.「Replace Fav」ボタン（コレクションは整理しましょう）**

**「My Favorite」シートの中で、フォーカスされているセルのプロンプトを、（「Favorite」タブの）テキストフィールドの内容と置換します（セルが空欄ならば、そのまま入力されます）。プロンプトが入っていないセル上でクリックしても、エラーが出ます。例えば、「Favに登録したプロンプトに、新しく要素を追加したら、もっとよくなった！　でも、お気に入りの中に似たものがダブるのは困る！」という場合に使えます。**

#### **5-2-9.「Export Fav」ボタン（性癖の遺産と、その保全）**

**お気に入りリストを、CSVファイルでエクスポートできます。「ナイスな絵が出た！　しかし、あの時のプロンプトって、何だったっけ！？」ということ、よくありますよね？　プロンプトは、ある意味「資産」ですから、それを保全するための機能です。**

##### **5-2-9-1.エクスポート（書き出し）の挙動**

**保存形式: 文字コード UTF-8 のCSV形式で出力されます。**

**ファイル名：デフォルトで「MyFavorite_yyyymmdd_HHmm（日付と時刻）」が設定されます。**

**上書き防止：保存先に同名のファイルが存在する場合、自動的に「(1)」「(2)」といった枝番が付与されます。既存のバックアップを誤って消去することはありません。**

#### **5-2-10「Import Fav」ボタン**

**CSVファイルから、お気に入りリストをインポートできます。ほら、仲間内で「俺はこんなすごいプロンプトを作ったぜ！」とか、交換したいじゃ名ですか？　そう言うときのための機能です。**

##### **5-2-10-1.インポート（取り込み）の挙動**

**取り込みモードの選択: 実行時に「既存のデータに追記」するか、「全クリアして新しく取り込む」かを選択できます。**

**データの自動分配:1つの行に2つ以上の値がある場合：1つ目を「Description（説明）」、2つ目を「Prompt（プロンプト）」へ振り分けます。**

**値が1つしかない場合：データの位置に関わらず、検索の利便性を優先して強制的に「Description」へ格納します。**

**スライド（詰め）機能: CSV内に形式不良や空行があっても、有効なデータだけを抽出して上方向へ隙間なく詰めてインポートします。**

**日本語のサポート: 日本語（全角文字）を含む説明も完全にサポートしています。**

**エラーログ: 制御文字などの不純物により取り込めなかった行がある場合、CSVと同じフォルダに詳細なエラーログ（ImportLog_時刻.txt）が生成されます。既存のリストに追加した結果が50件を超える場合、あふれたデータは、エラーログに記録されます。**

**注意：インポート元のファイルも、文字コードがUTF-8である必要があります。違っていた場合は、エラーが出て処理されません。**

**注意2：仕様上、「データのヘッダを飛ばして4行目から取得を開始する」挙動をします。それより上にあるデータは、取り込めません。**

#### **5-2-11.「Fav All Clear」ボタン**

**「My Favorite」シート内のデータを、全て消去します。元には戻せませんので、くれぐれもご注意を。**

## **6\. プロンプトの構築フロー**

**薄緑色のセルを左から順に選んでいくだけで、一つの完成されたプロンプトになります。**

**キャラ数(Character Count) → 肌の色・属性(Skin & Attributes) → 体型(Body Type) → 髪の長さ(Hair length) → 前髪(Bangs) → 髪の結び目(Tying) → 髪の色(Hair Color) → 職業(Occupation) → 下着(Underwear) → 服装(Outfit) → 服の状態(Outfit State) → ヘッドウェア(Headwear) → 足元周り(Footwear & Legwear) → アクセサリー類(Accessories) → 場所(Location) → 時間帯・周囲の状況(Time & Surroundings) → 体位(Position) → 行為・動作(Action & Movement) → 手段・道具(Means & Props) → 身体の部位(Body Parts) → 行為の状態(Interaction State) → 表情・生理現象(Expressions & Physical States) → その他アイテム(Misc Items) → アングル(Camera Angle) → 修正(Censorship Fixes)**

**メインシートの冒頭には、各項目のセルへ飛ぶハイパーリンクが設定しており、クリックするとジャンプし、フォーカスの移動したセルは、自動的にウィンドウの左端に寄ります。また、各項目の末尾には、「凡例に戻る」リンクを設置しています。**

## **7,「Sample Prompts」（サンプルプロンプト）シート**

**作者お気に入りのシチュエーションを収録しています。性癖の開示！（電波）サンプルをコピーすると、「Cockpit」タブウィンドウのテキストボックスにも同時に転送されるので、自分で調節して、オリジナルのプロンプトを作る事も出来ます。ただし、その場合の重み付けなどは、ご自身でお願いします。**

### **Tips：ネガティブプロンプトについて**

**縛りすぎはAIのバグを誘発します。最小限に留めるか、いっそオフにするのがコツです（「野郎の出演」等、絶許な結果を防ぐ程度に）。**

## **8.「作者覚え書き(ja) / Author's Notes (Tips and Rants)」シート**

**プロンプトに関するTipsとか、このマクロを作るに当たっての苦労話とかを書いています。息抜きにどうぞ（？）**

## **9.「CONTACT」シート**

**このREADMEにも書いていますが、作者の連絡先ブログなどを記載しています。**

## **10\. 免責事項・連絡先**

- **生成結果はAI次第です。当ツールの使用による損害について、作者は一切の責任を負いません。**
- **検証モデル：[RIN Anim8Draw Illustrious - Anime Drawing Model(Ver.4.0A)](https://www.seaart.ai/ja/models/detail/d158n1te878c73atvtdg)**
- **同・併用したLoRA：**[**Detailed anime style - Illustrious V2.0**](https://www.seaart.ai/ja/models/detail/33ff0599ab0e8f9b66d7bee70551df23)
- **作者：不二川巴人（ふじかわ ともひと）（「でぇすて」とか、「不二川“でぇすて”巴人」名義で、エロゲーライターをやっていました）**
- [**連絡先・ブログはこちら。**](https://dsblog.biz/)
- **リクエストや感想、あるいはバグレポートは、ブログのメールフォームまで。**[**投げ銭（PayPal）**](https://paypal.me/dst0508https:/paypal.me/dst0508)**も歓迎です！**
- **バグレポートももちろんですが、「こんなロケーションが抜けてるぜ！」とか、「この服がないぞ！」というフィードバックは、是非ともお寄せください。自分好みに自由に改変できるとは言え、「未知のシチュエーションを、俺も見たい！」からです！　あなたの意見が、このマクロをよりKENZENにします！**

**（おまけ）**

- [**SeaArtの個人ページ**](https://www.seaart.ai/ja/user/4b23d22e331a382c4adc23a3df4e7077?u_code=XWACJSXI)**では、生成したイラストを元に、書き下ろしショートショートを投稿したりしています。よろしければ、そちらもどうぞ。**

**さあ！　健やかなる()AIライフを！**
