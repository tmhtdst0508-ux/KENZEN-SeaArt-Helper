**[(日本語は下にあります)](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper#kenzen-seaart-helper-%E3%81%B8%E3%82%88%E3%81%86%E3%81%93%E3%81%9D)**

**\[Notice\]** This tool runs on **Excel Macro (VBA)**. It requires a **Windows/Mac PC environment** to unleash its full power. Sorry, prompt mages on mobile, this "forbidden door" requires a desktop key! 💻

  **\[注意\]** 本ツールは、**Excelマクロ(VBA)です。** 動作のためには、Windows/Mac PCが必要です。モバイル術士の皆様には申し訳ないのですが、「禁断の扉」を開けるには、「デスクトップ環境」というキーを入手してください。 💻
* * *
# Welcome to the KENZEN SeaArt Helper! 🤝

**Greetings to all fellow AI Mages in Japan and across the world who strive for "KENZEN" (Wholesome™/NSFW) art!**

I am Tomohito Fujikawa (aka "Dst" or "Deesute"). I am a writer who has spent approximately 15 years on the front lines of over 60 PC eroge(Visual Novel) titles, weaving scenarios and situations. In my pursuit of the ultimate "KENZEN (NSFW)" AI illustrations, I hit a wall: "I need a system to manage my deep-seated karma (fetishes) and complex prompts more intuitively and safely."

That is exactly why I forged this tool for my comrades (read: you degenerates). The KENZEN SeaArt Helper is the ultimate cockpit for AI mages wielding SeaArt and Stable Diffusion.

**How to Download**


![How to Download](images/How_to_Download_20260529.jpg)

**How to Unlock** `.xlsm` 


![How to Unlock xlsm](images/How_to_Unlock_xlsm_20260529.jpg)

**Screenshot**


![Screenshot](images/Screenshot_20260529.jpg)

 🚀 **【v2.4.0 Released!】** 🚀

**No More Auto-Resetting to Legend! (Pure QoL Update)**

Previously, the screen snapped back to the Legend every single time you copied a cell. But let’s be real—when you're cooking prompts, you often want to grab multiple words from the same column. Since the Cockpit already has a Category Jump ComboBox and a "Back to Legend" button, having it force-reset every time was honestly a massive pain in the ass and total eye strain. We killed that behavior entirely. Enjoy the smooth UX!

**Instant Prompt Sorter (Stop Forgetting Your Tags!)**

When you're building massive, complex prompt spells, it’s incredibly easy to forget to insert crucial tags. (Source: Trust me, I'm the dev and I do this daily). While AI can sometimes understand a messy prompt, it always performs better when the tags are beautifully ordered. That's why we added a feature that lets you sort the spells in the Cockpit’s main field with just a single click!

 **New "Wrap \[ \]" Feature for Multi-Character Cooking!**

Want two characters in one frame with different outfits and actions? I’ve been experimenting with the `BREAK` syntax myself, and noticed that enclosing each character's traits within `[ ]` dramatically increases your generation gacha win-rate. So, I built it right in! Along with this, `BREAK` has been officially added to the database, ensuring clean line breaks before and after whenever you copy it.

**LoRA Overhaul: Space Separation & Individual Weights!**

After the v2.3.0 release, I did some deep diving into the prompt meta. Turns out, forcing LoRA tags to the very front isn't necessary, and using commas to link them is actually an anti-pattern. So, I abolished the old "insert LoRA at the front" behavior. LoRAs are now cleanly linked with standard half-width spaces instead of commas. Furthermore, when stacking multiple LoRAs, you are no longer forced to apply the same uniform weight—you can now customize the weight for each trigger word individually!

**"Surprise Me!" – Let Gemini Take the Wheel!**

Pure gacha mode for fun! We already have the god-defying feature of letting Google Gemini engineer NSFW prompts for you, but I thought: "What if we just let the AI go wild?" Check this option, and Gemini will craft a mind-blowing prompt based on words randomly pulled from your database.

**Civitai-Friendly Prompt Import & Cleaning!**

Look, my bad on this one. When browsing Civitai, model descriptions often come with recommended positive and negative prompts already heavily weighted. In previous versions, copy-pasting them directly would import the raw `()` brackets, messing up your database records. I fixed that! Plus, the Positive Prompt management screen now supports bulk copy-paste registration straight from your browser.

**Dynamic Prompt & Wildcard Support!**

For the local Stable Diffusion Web UI gigachads: We’ve bundled wildcard files compatible with the "Dynamic Prompts" extension inside the release ZIP! These cover hairstyles, facial expressions, and camera angles. A dedicated "Dynamic Prompt Wildcard" category has also been added to the database.

**Other Bug Fixes**

Forgive me, standard dogeza applied. m(_ \_)m

If this tool aids you on your glorious creative journey, hitting that ⭐Star button on this repository would be the ultimate encouragement for the author! Let us spread the "KENZEN" culture together!

# ■ KENZEN SeaArt Helper Manual (v2.4.0)

_Prefer offline reading? [Download the PDF Manual here!](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper/blob/main/KENZEN_SeaArt_Helper_Manual_v.2.4.0.pdf)_

**TL;DR:** A specialized tool for building NSFW generation prompts for SeaArt (& Stable Diffusion). Using an API key, you can also make Gemini brainwash... I mean, brainstorm NSFW prompts for you.

# 1\. Introduction

This Excel workbook utilizes macros.

- **Coding Assistance:** Google Gemini (A.K.A. the guy I dumped all the actual implementation work on).
- **Transparency:** Press `Alt+F11` to open the VBA editor and view/modify the MainCode inside the \[Main Module\].
- (_Note: I haven't planted any malicious backdoors, but it does contain spaghetti code boiled exactly to Google Gemini's instructions._)
- **Copyright:** I am not abandoning my copyright, but you are free to modify and redistribute it. No need to report to me (though it makes me happy if you do).

# 2\. Purpose of Creation

Developed to dramatically streamline the creation of "NSFW-specialized" prompts in SeaArt. It is an English prompt-building tool designed to accurately convey your intentions to the AI. I packed it with my passion and fetishes, driven by the thought: "I want something easy for _me_ to use!" Avoid the pitfalls of "broken English" and take the shortest route to your ideal output. ...It's completely! KENZEN (Wholesome), I swear!

# 3\. Key Macro Features (Safety Design)

## 3-1. Full-Width (Double-Byte) Check

If Japanese (full-width characters) is included, it halts the copy process and alerts you with a warning sound.

## 3-2. Auto-Concatenation

From the second copy onward, it automatically inserts ", " (comma + half-width space) and appends the text.

## 3-3. Real-Time Display

Your prompt-in-progress is constantly displayed in the "Current Prompt" field.

## 3-4. Reset Function

The Clear button wipes both the clipboard and the screen display.

# 4\. Preparation for Use

## 4-1. Unblock

Right-click the downloaded .xlsm file → Properties → Check "Unblock" (or "Allow") and save.

## 4-2. Enable Macros

Open the file and click "Enable Content" at the top.

## 5\. Search Panel & Operation Guide

Upon launching (enabling macros), the main window will open.

Even if you close it with the top-right "X", you can reopen it anytime via "Open Main Window" or the shortcut key `Ctrl + Shift + O`. You can also minimize it and banish it to the bottom-left corner of your desktop. Because honestly, the window gets in the way when you're looking at a freaking massive database list, right?

The window consists of 6 tabs (panels): "Cockpit", "Positive", "Negative", "LoRA", "Favorite", and "Gacha!".

### 5-1. "Cockpit" Tab

The tab containing menus for building your prompt.

### 5-1-1. "Current Prompt"

- Displays the current contents of your clipboard. Every time you copy a prompt, it appends here. Since it's a text field, manual fine-tuning is totally possible.

### 5-1-2. Search Box (Enter Keyword)

- Enter a keyword and hit Enter or click "Search" to search the main sheet. When a hit occurs, the cell glows yellow, and the focus moves to the prompt cell to its right. If you search in English, it stays exactly on that spot. For multiple hits, pressing Enter loops to the next result.

### 5-1-3. "Jump to Category" Combo Box, "Stay Here" Checkbox, "Back to Legend" Button

- Use the dropdown to jump to category header cells. As explained later, copying a cell usually returns you to the "Legend (凡例)" cell. But what if you want multiple words from the same category? (For example, the golden combo of "Huge Breasts", "Huge Ass", and "Thick Thighs"?). Returning to the Legend every time is annoying. In that case, check "Stay Here". Copying a cell won't return you to the Legend. Also, there's a "Back to Legend" hyperlink at the end of each item, but scrolling all the way down is a pain! Just click the "Back to Legend" button instead.

### 5-1-4. "Copy" Button

 - Clicking on a prompt cell instantly transfers it to your clipboard. As you keep copying, words accumulate in the "Current Prompt" field, automatically joined by a comma and a space (`,` ). The `BREAK` tag is handled with special care, inserting clean line breaks before and after.

### 5-1-5. "Copy without comma" Button

- Concatenates without a comma (just a half-width space). Useful when combining words into a single concept, like oversized tank top. The macro's behavior is: "Except for the very first word, add a comma and space to the beginning of the copied word." So, for example, for `"1 girl"` `"is going to"` `"the park"` `"with"` `"me"`: Normal copy `"1 girl"`, then comma-less copy `"is going to"`, comma-less copy `"the park",` comma-less copy `"with".`.. keep doing comma-less copies until the spot where you actually _want_ a comma (in this case, before `"me"`). If you normally copy the next prompt, it will add the ", " separator.
-  *Note*: Clicking this for the very first prompt will _not_ add a leading space.

### 5-1-6. "Undo" Button

- Reverts your action. Up to 50 times. Clipboard contents are also reverted.

### 5-1-7. "Clear" Button

- Erases the clipboard and the "Current Prompt" field. Undo is possible.

### 5-1-8. "All Clear" Button

- Completely erases the clipboard, "Current Prompt" field, and ALL history. Resets everything. Undo is NOT possible. (A.K.A: The Nuke / Physical Scorched Earth Button).

### 5-1-9. Weighting Feature ("Weight" Checkbox & "Wrap Block" Button)

- You can weight prompts you want to emphasize (or weaken) from 0.5 to 1.3. Not just single words, but combinations like the aforementioned oversized tank top can become (oversized tank top:1.1). Clicking "Wrap Block" encloses the block up to the preceding comma in () and applies the weight. Naturally, this works for single prompts too. To prevent misclicks and reduce AI load, the checkbox turns off once you click "Wrap Block". Also, if you turn the check back on without adding a new prompt and click "Wrap Block", it removes the weighting.

### 5-1-10. "Wrap \[ \]" Button

When using the `BREAK` syntax, wrapping the entirety of each character's traits in `[ ]` helps the AI understand the prompt layers much better (depending on the model you use). Simply select a block of text inside the textbox and click this button to instantly enclose it in `[ ]`.

### 5-1-11. "Done!" Button

- Once you're satisfied (or tired) of manual tweaking, click "Done!" to copy the current contents of the text field.

### 5-1-12. "Everyone, Fall in! (Sort Prompt)" Button

 - Building long prompt chains often leads to "forgotten tags." While the AI can read disorganized prompts, keeping them tidy yields the best results. Click this button to instantly sort your prompts according to the order of the database. *Note: If you use the BREAK syntax, this will only apply to the final paragraph.*

### 5-1-13. Send to Fav Button

- Sends your favorite prompts to the text field in the "Favorite" tab window.

## 5-2. "Positive" Tab (AI Only Does What It's Told)

A dedicated screen to efficiently manage and deploy positive prompts to make the fundamentally lazy AI actually do its job.

### 5-2-1. Your Positive Stock

- The list box on the left. Displays your currently registered positive prompts. Defaults are set upon initialization. You add, subtract, divide, and multiply (?) from here. Multiple selection is possible by holding Ctrl or Shift.

#### 5-2-1-1. Selected Delete Button

- Deletes selected items from your stock. The original data is also deleted.

#### 5-2-1-2. Delete All Button

- Deletes ALL stock. Be careful, all original data will vanish.

#### 5-2-2. Applied Positive List

- The list box on the right. This is where the elite troops (?) you intend to output as positive prompts gather. Also allows multiple selection.

#### 5-2-2-1. Dismiss All Button

- Rejects (erases) all words you intended to apply. Does not affect original data.

#### 5-2-3. ▲/▼ Buttons

- Buttons next to the two list boxes. As you can see, they sort the order in the list. However, to prevent errors, you can only move one at a time. The left stock changes the original data order, but the right side does not.

#### 5-2-4. →/← Buttons

- Also obvious, but "→" sends selected items from stock to the applied list, and "←" returns (dismiss) items you were going to apply. Does not affect original data.

#### 5-2-5. Add Button

- Sends the selected tags to the "Preview" text box at the bottom. If multiple are selected, they output separated by ", ".

#### 5-2-6. Set as Default Button

- Registers the contents of the Preview window as your default positive prompt.

#### 5-2-7. Save as Preset Button

- Registers the contents of the Preview window as a preset.

#### 5-2-8. Add New Positive Stock? Textbox / Add New Button / Clear Input Button

- To register a new positive prompt, enter it here and hit Enter or "Add New" to add it to the left stock list. Commas, tabs, or line breaks will be formatted and registered correctly. Clear Input simply clears the textbox.

#### 5-2-9. Preview Box

- As mentioned, the actual prompt is displayed here. Direct manual editing is also possible.

#### 5-2-10. Send to Cockpit Button / Clear Positive Button

- Transfers the positive prompt shown in Preview to the Cockpit's main text window. You can also do this with the shortcut `Ctrl+Shift+P`. It will throw an error if the Preview is empty. Clear Positive clears the Preview window.

#### 5-2-11. Positive Preset Combo Box / Call Preset / Delete Preset

- Select registered presets from the dropdown and use "Call Preset" to summon them. You can delete them with "Delete Preset", but the (Default) cannot be deleted.

### 5-3. "Negative" Tab (The "Absolutely Unforgivable" Blacklist)

A dedicated screen to efficiently manage and deploy negative prompts to stop the AI from "going rogue." It's basically symmetrical to the Positive prompt management screen, but with minor differences.

#### 5-3-1. Your Negative Stock

Left list box. Same as Positive. Defaults are set.

##### 5-3-1-1. Selected Delete Button

Deletes selected items. Original data gone.

##### 5-3-1-2. Delete All Button

Same as Positive. All original data goes boom.

#### 5-3-2. Applied Negative List:
Same as Positive.

##### 5-3-2-1. Dismiss All Button

Same as Positive (Ditto).

#### 5-3-3. ▲/▼ Buttons

Same as (omitted).

#### 5-4-4. →/← Buttons

Posi- (omitted).

#### 5-4-5. Add Button

Po- (omitted).

#### 5-4-6. Weighting Checkbox, Combo Box, Weighten and Add Button

Just like the Cockpit screen, checking the box opens the combo box, allowing you to weight negative prompts. Again, from 0.5 to 1.3.

#### 5-4-7. Add New Negative Stock? Textbox / Add New / Clear Input

Don't make me explain the same thing twice! (Sudden irrational anger).

#### 5-4-8. Preview Box

Needs no explanation.

#### 5-4-9. Copy Negative / Clear Negative Button

Copies the negative prompt to the clipboard, or clears the Preview box.

### 5-5. "LoRA" Tab (The LoRA Forge)

Visually and intuitively build and manage combinations of the essential &lt;`lora:hash:strength`&gt; format prompts and their accompanying trigger words required for quality enhancement.

_(If no LoRA is registered, the UI is locked for safety. Please register LoRAs from "Open LoRA Manage Window".)_

#### 5-5-1. Use LoRA Checkbox

Toggles the features of this tab on/off.

#### 5-5-2. Open LoRA Manage Window Button

Opens the dedicated window (described later) to register and manage LoRA system hashes and trigger words.

#### 5-5-3. LoRA Selection & Strength

Select a registered LoRA from the dropdown. Once selected, the recommended Strength and registered trigger words automatically expand.

#### 5-5-4. Trigger Words Checkbox

Check only the trigger words you want to use this time from the expanded list (User-friendly design: they all default to ON). LoRAs without trigger words won't show anything here.

#### 5-5-5. Set! & Cancel Buttons

"Set!" puts the selected LoRA and trigger words into the right list box (Cart). "Cancel" wipes the selection clean.

#### 5-5-6. Your Selected LoRA (Cart)

List of currently set LoRAs. To stack multiple, select another LoRA and press "Set!" to keep adding.

#### 5-5-7. Wrap LoRA! Button & Weight (Prompt Alchemy)

Formats the trigger words of the LoRA currently in your cart into the `<lora:Hash:strength>, (Trigger word:1.1)` format and outputs it to the Preview field. If the trigger word weight is exactly 1.0, the `()` brackets are automatically omitted. *Note: If you want to apply different weights to multiple trigger words within the same LoRA, the current UI doesn't support it out of the box—please manually edit the Preview textbox.*
_Safety Feature:_ In the unlikely event a "Ghost LoRA" (deleted from the main management data) remains in the cart, pressing this button automatically detects and exorcises (deletes) it from the cart.

#### 5-5-8. Remove / Forget LoRA Buttons

"Remove" kicks one selected LoRA out of the list. "Forget" completely nukes the cart and preview so you can start over.

#### 5-5-9. Send to Cockpit / Send to Fav Buttons

Transfers the completed LoRA prompt in Preview to the text field of the Cockpit or Favorites tab. If there's already a prompt in the Cockpit, it inserts it at the very beginning.

#### 5-5-10. Preset (Save & Call Presets)

"Save as Preset" lets you name and save the current LoRA combination (cart contents and weights). "Call Preset" summons it instantly anytime, and "Delete Preset" deletes it. Saving frequent outfit or art style combos will drastically cut your time.

### 5-6. "LoRA Register & Manage" Window(The LoRA Vault)

The beating heart where you register and manage actual LoRA data (hashes and trigger words). Accessed via "Open LoRA Manage Window" on the LoRA tab.

#### 5-6-1. LoRA Alias / System Hash / Recommended Strength

**Alias:** Any name you can easily identify (e.g., JK Uniform, Watercolor Style).

**System Hash:** The actual hash string used by SeaArt, etc.

_Note: What is a Hash?_ When you open a specific LoRA's detail page on SeaArt, look at the URL: https://www.seaart.ai/models/detail/40095be8759dde4285ccf683b24e8852.
The string after /detail/ (40095be8759dde4285ccf683b24e8852) is the hash.

If you're using Stable Diffusion in a local environment, it's the filename of the `.safetensors` file you have.

**Recommended Strength:** The recommended value for this LoRA's intensity.

_Auto-Sanitize:_ Even if you paste full-width characters or invalid symbols into the Hash field, an ironclad bouncer is on duty to automatically purify it into safe half-width characters.

#### 5-6-2. LoRA Trigger Word / None Checkbox

Enter the trigger word required to activate the LoRA and click "Add Trigger". For multiple words, comma-separation works.

_Note:_ Max 10 trigger words. It runs a case-insensitive duplicate check, and harmful symbols for prompts (! ? @ # ( ) etc.) are blocked with an error. If a LoRA doesn't need a trigger word, check "None".

#### 5-6-3. Register / Cancel / Update LoRA Buttons

Registers your input. While editing an existing LoRA (Manage)—like adding or removing trigger words—the button morphs into "Update LoRA". A dirty-check function prevents pointless overwriting if nothing changed.

#### 5-6-4. Registered LoRA List & Sort Buttons (▲ / ▼)

List of registered LoRAs. Use the "▲ (Up)" and "▼ (Down)" buttons to push your favorite LoRAs to the top or sort by category. This sort order directly syncs with the dropdown on the main screen.

#### 5-6-5. Manage / Delete LoRA Buttons

Edit (Manage) or completely obliterate (Delete) the selected LoRA from the list.

#### 5-6-6. Export / Import LoRA (CSV Backup):

**Export:** Dumps your precious LoRA collection into a UTF-8 CSV file. Auto-appends a branch number if the filename duplicates.

**Import:** Reads LoRA data from a CSV. Capable of "appending" to existing data, featuring a smart merge function that automatically skips already registered LoRAs (where Alias and Hash match perfectly).

### 5-7. "My Favorite" Tab (The Vault of Fetishes)

The tab holding the menus to manage the precious prompts you crafted through blood, sweat, and tears (?). Note: All buttons here only work if the "My Favorite" tab is active.

#### 5-7-1. Search Fav Field, Search Fav / Clear Search Buttons

Enter a keyword and hit Enter or "Search Fav" to search within the "My Favorite" sheet. It only searches the "Description" field. Just like the main sheet, hit cells glow, and it loops on multiple hits. Clear Search simply empties the search box.

#### 5-7-2. "Pull From Cockpit" Button

Transcribes the prompt currently in the Cockpit text field.

#### 5-7-3. "Send to Cockpit" Button

The reverse of the above. Sends the Favorite Prompt text field contents to the Cockpit. Great for refining a favorite by adding main sheet prompts to it.

#### 5-7-4. "Clear Fav" Button

Clears all fields in the "Favorite" tab.

#### 5-7-5. "Copy Fav" Button

Copies a favorite from the "My Favorite" sheet. I mean, you definitely want to reuse a spell you like, right? The copied spell (prompt) appears in the text field along with its description. (Nobody can remember what a spell does just by looking at the raw prompt without a description, right?)

#### 5-7-6. "Tweaked!" Button

Becomes clickable after manually tweaking a copied favorite prompt inside the text field. Clicking it copies the new text field contents.

#### 5-7-7. "Add to Fav!" Button

Registers the prompt in the field to the "My Favorite" sheet. Throws an error if "Description" is blank. Max 50 slots.

#### 5-7-8. "Replace Fav" Button (Keep your collection tidy)

Replaces the prompt in the currently focused cell of the "My Favorite" sheet with the contents of the text field (if the cell is blank, it just inputs it). Clicking on a cell with no prompt throws an error. Useful for: "I added elements to a Fav prompt and made it way better! But I don't want duplicate/similar prompts cluttering my favorites!"

#### 5-7-9. "Export Fav" Button (Legacy of Fetishes & Preservation)

Exports your favorite list as a CSV. You know the feeling: "This output is God-tier! Wait, what was the prompt again!?" Prompts are, in a sense, "assets," and this feature preserves them.

##### 5-7-9-1. Export Behavior:

Format: UTF-8 CSV.

Filename: Defaults to `MyFavorite_yyyymmdd_HHmm`.

Overwrite Protection: Auto-appends (1), (2) to prevent wiping out existing backups.

#### 5-7-10. "Import Fav" Button

Imports a favorite list from a CSV. Because you totally want to flex to your buddies: "Look at this insane prompt I made!" and trade them, right?

##### 5-7-10-1. Import Behavior:

Mode Selection: Choose between "Append to existing data" or "Clear all and import fresh".

Auto-Distribution: If a row has 2+ values: 1st goes to "Description", 2nd to "Prompt". If only 1 value: forced into "Description" for search convenience.

***Slide (Pack) Feature***: Skips bad formats/blank rows and packs valid data upwards without gaps.

***Japanese Support***: Fully supports descriptions containing Japanese (full-width).

Error Log: Generates ImportLog_Time.txt for rows containing impurities like control characters. Also logs overflow data if appending exceeds the 50-limit max.

_**Note 1:**_ The source CSV MUST be UTF-8. Otherwise, error.

_**Note 2:**_ By design, it skips headers and starts reading from Row 4. Anything above that is ignored.

#### 5-7-11. "All Clear Fav" Button

Wipes ALL data in the "My Favorite" sheet. Cannot be undone, so use with extreme caution.

### 5-8. "Gacha!" Tab (Leave it to AI / Prompt Alchemy)

"If you get tired of thinking up spells yourself, just dump it on Gemini, **can't you?** **_\[Menacing ゴゴゴゴ\]_**" (in a JOJO style). Even I, the author, think, "How the hell did I pull this off!?" It's an insanely (dead slang) good feature.

In short, it's an automatic prompt generation feature utilizing the Google Gemini API. When you are absolutely sick of building dense, complex spells (prompts), surrender to the AI's imagination as a "pure breather." Literally a "Gacha" of "What will pop out?". By the way, I've cast a certain magic (exaggeration) on Gemini, so it _will_ expertly alchemize NSFW prompts. However, there are some instant-death landmine words. See the "Author's Notes (Tips and Rants)" sheet for details.

#### 5-8-1. Google Gemini API Key

You must acquire your own API key from Google AI Studio and input it to use this.

- _**No Support Provided:**_ I will NOT provide any individual support on how to get or set up API keys. Only for "Mages" who can solve it themselves. Let’s be real here—I’m a writer, not your personal IT support!
- _**AI's Mood:**_ Google Gemini is extremely "well-mannered." Even though I say it can output NSFW prompts, depending on the content and the AI's "mood" (safety filters occasionally kicking in hard), it might fail or reject you. Please enjoy that as part of the "Gacha" experience.
- _**Auto-Save API Key:**_ Once Gemini successfully alchemizes a prompt, that success acts as a flag. The API key is saved to your local machine's registry and auto-fills next time. If you change your key, just overwrite it; it saves on the next successful run.

#### 5-8-2. Natural Language Input Field (Please tell me your desire!)

Freely type the situation (desire) you envision using normal words.

- _**Bilingual Support:**_ Japanese, English, or a mix—the AI understands and sublimates it into the optimal tagset for generative AIs like SeaArt.

#### 5-8-3. Operation Buttons

- **Feeling Lucky?:** Sends your input to Gemini and rolls the gacha. Error responses _do_ consume API calls, so spamming it just wastes your ammo. Thus, it grays out for 15 seconds after one click. During this cooldown, you are locked out of moving to other tabs to prevent misclicks. It's a feature, not a bug (ancient proverb).
- **Clear:** Clears the input field.

#### 5-8-4. Generation Result Area (How about this?)

Displays the AI-alchemized prompt.

- **Copy This!:** Copies it to the clipboard.
- **Send to Cockpit / Send to Fav:** Sends it to your preferred tab. Apply your favorite LoRAs or whatever you please.
- **Clear Result:** Clears the result.

#### 5-8-5. "Trigger Happy?" Frame (Gacha Ammo Remaining)

Shows your remaining daily gacha rolls.

- _**Rough Estimate:**_ This number is just a "guide" based on typical Google free-tier limits and internal counting. Actual limits may vary.
- _**Reset:**_ Resets at 0:00 Pacific Time (PT).
- _**Limit Adjustment (Hidden Feature):**_ For heavy users paying for the API who want to change the cap, double-click the counter number. A box will appear to freely adjust your daily limit.

### 5-8-6. "Surprise Me!" Checkbox / SFW, NSFW & Hardcore Options

Checking this box reveals three option buttons. Select your desired vibe and hit "Feeling Lucky?" to let Google Gemini engineer a creative prompt based on words randomly selected from your database.

_**Tips: Gacha Tricks:**_ When speaking your desires, _you_ don't need to be "well-mannered." During debugging, I found that if you show weird hesitation, Gemini senses it and triggers its (cheeky) safety filters. That's right. If you mean female genitalia, don't sugarcoat it with "crotch." Just write "p\*ssy" straight up.

## 6\. Prompt Building Flow

Simply select the light green cells from left to right to complete a single prompt.

Character Count → Character Placement → Skin & Attributes → Body Type → Wildcard for Hair → Hair length → Bangs → Tying → Hair Color → Body Hair → Occupation → Underwear → Outfit → Outfit State → Headwear → Footwear & Legwear → Accessories → Location → Time & Surroundings → Position → Action & Movement → Bondage Action & Movement→ Means & Props → Body Parts → Interaction State → Expressions → Body fluids → Misc Items → Camera Angle → Censorship Fix, Others

If you are using the Dynamic Prompt extension in a local Stable Diffusion environment, I have prepared wildcards for hair,expressions, and angles. These are included in the ZIP file, so please feel free to use them.

At the top of the main sheet, there are hyperlinks to jump to each category cell. Clicking one automatically aligns the focused cell to the left edge of the window.

At the top of the main sheet, there are hyperlinks to jump to each category cell. Clicking one automatically aligns the focused cell to the left edge of the window.

## 7\. "Sample Prompts" Sheet

Contains the author's favorite situations. Full disclosure of fetishes! (Static noise). Copying a sample also transfers it to the Cockpit text box, so you can tweak it to create an original prompt. However, you're on your own for setting weights.

## 8\. "Author's Notes (Tips and Rants)" Sheet

Contains prompt tips and tales of my struggles while making this macro. Read it when you need a break (?).

## 9\. "CONTACT" Sheet

As written in this README, contains the author's contact blog, etc.

## 10\. Disclaimer & Contact

Generation results are entirely up to the AI. The author assumes ZERO responsibility for any damages caused by using this tool. Likewise, the "Gacha!" feature does not guarantee a successful NSFW output.AI results are unpredictable. I am not responsible for what you generate.

- Tested Model: - [REED_XXX_illustrious_SDXL V14.0](https://civitai.red/models/1717562/reedxxxillustrioussdxl)
- Contact/Blog: [dsblog.biz](https://dsblog.biz/)
- Support the Dev: Bug reports are great, but [PayPal tips](https://paypal.me/dst0508) keep the lights on!
- While bug reports are absolutely welcome, what I _really_ crave are your missing tag requests! Hit me with feedback like, "Hey, you forgot this location!" or "Where the hell is this specific outfit?!" Sure, you have the freedom to mod the code and add them yourself, but please share them with me—because _I too wish to gaze upon uncharted scenarios!_ Your collective wisdom is the fuel that makes this macro even more totally KENZEN (Wholesome™)!
- Author: Tomohito Fujikawa (aka “Dst” or "Deeste" / Former Eroge (Visual Novels) Writer).
- Bonus: Craving some Lore? Check out [my SeaArt page!](https://www.seaart.ai/new-user/4b23d22e331a382c4adc23a3df4e7077) I put my VN writing skills to work by posting original short stories alongside my generated art.

**ENJOY your KENZEN AI Life!** 😊

* * *
<a id = "kenzen-seaart-helper-%E3%81%B8%E3%82%88%E3%81%86%E3%81%93%E3%81%9D"></a>
# KENZEN SeaArt Helper へようこそ！🤝

**日本の、そして世界中の「KENZEN」なる同志たるAI術師の皆様、ようこそ！**

不二川巴人（ふじかわ ともひと。あるいは「でぇすて」）と申します。約15年間にわたり、60タイトル以上のPC美少女ゲームの最前線で、シナリオやシチュエーションを紡いできた物書きです。究極の「KENZEN（NSFW）」なAIイラストを追い求める中で、一つの壁にぶち当たりました。それは、「己の深淵なる業（フェティシズム）と複雑怪奇なプロンプトを、もっと直感的に、かつ安全に管理するシステムが必要だ」ということです。

だからこそ、同志たち（と書いて「お前等」と読む）のためにこのツールを錬成しました。**KENZEN SeaArt Helper**は、SeaArtとStable Diffusionを駆使するAI術師のための究極のコックピットです。

**ダウンロード方法**


![How to Download](images/How_to_Download_20260529.jpg)

 `.xlsm`**ファイルのロック解除方法**

 
![How to Unlock xlsm](images/How_to_Unlock_xlsm_20260529.jpg)

**スクリーンショット**


![Screenshot](images/Screenshot_20260529.jpg)

🚀 **【v2.4.0 リリース！】** 🚀

**コピー後の挙動変更！**
今までは、セルをコピーするたびに、凡例へ戻っていました。しかし、同一の列から複数の単語をコピーしたいことも多々あります。Cockpitにカテゴリジャンプのコンボボックスと「Back to Legend」ボタンがありますから、いちいち戻っていては、目が滑ってぶっちゃけうっとうしい！　なので、戻る挙動そのものを廃止し、UXの向上を図りました。

**プロンプトソート機能搭載！**
長大な呪文を構築していると、「タグ（要素）の入れ忘れ」というケースが、頻繁に発生します。ソースは作者です。順番が乱れていても、いいと言えばそうなのですが、AIにより正しく理解させるためには、やはり順番通りであった方がいい。そんなわけで、Cockpit内のメインフィールドの呪文を、ボタン一発でソートする機能を実装しました。

**「Wrap[ ]」機能追加！**
「画面内に人物が2人いる場合、それぞれに違う服装や行動をさせたい」という状況のために、作者も「BREAK」構文を覚え、その際、各人物にまつわる要素全体を、”[ ]”で括った方が、打率が上がることに気付きました。じゃあ追加してやれ！　ってことで。それに関連して、データベースにも「BREAK」を追加し、それをコピーしたときには、前後に改行が入るようにしました。

**LoRA関連のUX変更！**
v2.3.0のリリース後、作者も色々調べました。それによると、LoRAに関するタグは、必ずしも先頭に挿入する必要はなく、むしろカンマで連結する方が非推奨らしい。じゃあってことで、前バージョンまでの、「LoRAのタグを先頭に挿入する」挙動を廃止し、カンマの代わりに半角スペースで連結するようにしました。また、LoRAを重ねがけする際、これまでは、選択したLoRAのトリガーワードに重み付けをする際、全てのLoRAに同じ重み付けがされていた挙動を改善し、個別に設定できるようにしました。

**「Surprise Me!」機能追加！**
純粋なお遊び機能です。「Google GeminiにNSFW絵のプロンプトを考えさせる」という、神をも恐れぬ機能を搭載しているわけですが、「任せっきりにしたら、面白いんじゃないか？」と思って、追加しました。データベース内からランダムに抽出された単語を元に、Google Geminiがプロンプトを考えてくれます。

**ポジティブ＆ネガティブプロンプト管理機能強化！**
これはもう、作者が完全に悪いのですが、Civitaiなどを見ていると、各モデルの概要欄に、推奨されるポジティブ＆ネガティブプロンプトが、既に重み付けされている状態で記載されていることが、多々あります。前バージョンまでは、そのまま（カッコ付きのまま）コピペすると、「`()`」等が入ったままの、誤った形で記録されてしまっていました。そこを直したのと、ポジティブプロンプトの管理画面についても、ブラウザからの一括コピペ登録ができるようにしました。

**Dynamic Prompt対応！**
Stable Diffusion Web UIのローカル環境限定機能ですが、拡張機能の「Dynamic Prompt」に対応したワイルドカードファイルを、配布ZIPの中に同梱しました。具体的には、ヘアセット、表情、アングルに関するものです。それに併せて、データベースにも、「Dynamic Prompt用ワイルドカード」の項目を追加しました。

**その他バグフィクス**
なんか、ごめん（土下座）


**もしこのツールが、あなたの素晴らしき創作の旅の役に立ったなら、このリポジトリに ⭐Star を押していただけると、作者にとってこの上ない励みになります！** 共に「KENZEN」な文化を広めていきましょう！

# ■KENZEN SeaArt Helper マニュアル（v2.4.0）

オフラインマニュアルは、[こちらをご覧下さい](https://github.com/tmhtdst0508-ux/KENZEN-SeaArt-Helper/blob/main/KENZEN_SeaArt_Helper_Manual_v.2.4.0.pdf)

**要約：SeaArt（＆Stable Diffusion）での、NSFW絵の生成プロンプト構築に特化したツールです。APIキーを使って、GeminiにNSFW絵のプロンプトを考えさせることもできます。**

## 1.はじめに

当Excelブックにはマクロを使用しております。

コーディング支援： Google Gemini（実質的な実装担当という名の丸投げ）

透明性の確保： Alt+F11 でVBAエディタを開き、\[標準モジュール\] 内の MainCode を参照・改変いただけます。（※悪意あるバックドアは仕込んでいませんが、Google Geminiの言いなりで煮込まれたスパゲッティコードが内包されています）

著作権： 放棄しませんが、改変および再配布は自由です。報告も不要です（あると作者が喜びます）。

## 2．作成目的

SeaArtでの「NSFW絵に特化した」プロンプト作成を劇的に効率化するために開発しました。 AIへの意図を正確に伝えるための、英語プロンプト構築ツールです。「自分用に使いやすいものを！」という情熱と性癖を詰め込みました。「ダメ英語」の落とし穴を回避しつつ、最短ルートで理想の出力を目指します。……全き！　健全（KENZEN）ですよ！

## 3.マクロの主要機能（安全設計）

### 3-1全角チェック

日本語（全角文字）が含まれている場合、コピーを停止し警告音で知らせます。

### 3-2.自動連結

2回目以降のコピー時、自動的に “, “（カンマ＋半角スペース）を挿入して追記します。

### 3-3.リアルタイム表示

構築中のプロンプトは「Current Prompt」フィールドに常時表示されます。

### 3-4.リセット機能

Clear ボタンでクリップボードと画面表示内容を消去します。

## 4．使用準備

### 4-1ブロック解除

ダウンロードしたxlsmファイルを右クリック → プロパティ→ 「許可する」にチェックを入れて保存。

### 4-2マクロ有効化

ファイルを開き、上部の「コンテンツの有効化」をクリック。

## 5．検索パネルと操作方法

起動する（マクロを有効化する）と、メインウィンドウが開きます。

右上の×で閉じてしまっても、Open Main Window、あるいは、ショートカットキー `Ctrl + Shift + O` から、いつでも開けます。また、最小化して、デスクトップの左下隅に追いやる（？）こともできます。だって、ヴァチクソに広大なデータベースの一覧見る時に、ウィンドウが邪魔でしょう？

ウィンドウは、「Cockpit」、「Positive」、「Negative」、「LoRA」、「Favorite」、「Gacha!」の6つのタブ（パネル）で構成されています。

### 5-1「Cockpit」タブ

プロンプトを構築していくためのメニューがあるタブです。

#### 5-1-1.「Current Prompt」

ここに、現在のクリップボードの内容が表示されます。プロンプトをコピーするたびに、追記されていきます。テキストフィールドなので、手動での微調整も可能です。

#### 5-1-2．検索ボックス（Enter Keyword）

フィールドにキーワードを入力し、エンターキー、ないしは「Search」ボタンをクリックすると、メインシート内を検索できます。ヒットすると、セルが黄色く光り、右隣のプロンプトのセルに移動します。英語で検索した場合は、その場にとどまり、移動しません。連続ヒットする場合、Enterを押すごとに次の結果へループ移動します。

#### 5-1-3.「Jump to Category」コンボボックス、「Stay Here」チェックボックス、「Back to Legend」ボタン

プルダウンで、各カテゴリの見出しセルへ飛びます。後述しますが、各セルをコピーすると、「凡例（Legend）」セルへ戻るのですが、「同じカテゴリから、複数の単語を選びたいとき」（例えば、「巨乳」と「巨尻」、「厚い太もも」の、黄金セット（？）とか）に、いちいち凡例まで戻るのはうっとうしい。そんな時は、「Stay Here」にチェックを入れてください。セルをコピーしても、凡例に戻りません。また、項目の最後に、「凡例に戻る」のハイパーリンクがあるのですが、これも、下までスクロールするのも面倒だ！　というときに、「Back to Legend」を押すと、凡例に戻ります。

#### 5-1-4.「Copy」ボタン

コピーしたいプロンプトのセル上でクリックすると、クリップボードに転送されます。コピーを繰り返すと「Current Prompt」フィールドに、「, 」（カンマと半角スペース）を付けて、プロンプトが蓄積されます。「`BREAK`」に関しては特殊で、前後に改行が入ります。

#### 5-1-5.「Copy without comma」ボタン

カンマなし（半角スペースのみ）で連結します。oversized tank top 等、連結して一つの概念を指す場合に有効です。マクロは、「最初を除き、コピーされる単語の頭に、カンマと半角スペースを付ける」挙動をします。なので、例えば、「`”1 girl”` `“is going to”` `“the park”` `“with”`  `“me”`」の場合は、まず、「`1 girl`」で通常コピー、次に「`is going to`」でカンマなしコピー、次の「`park`」でカンマなしコピー、その次の「`with`」でも、カンマなしコピー……と、「次にカンマを入れるべき所」（例の場合は「`me`」）まで、カンマなしコピーをしてください。次のプロンプトを通常コピーすれば、区切りに「, 」が付きます。なお、最初のプロンプトを、このボタンでクリックしても、先頭にスペースは入りません。

#### 5-1-6.「Undo」ボタン

操作を戻せます。50回まで可能です。クリップボードの内容も、元に戻ります。

#### 5-1-7.「Clear」ボタン

クリップボードと「Current Prompt」フィールドを消去します。アンドゥは可能です。

#### 5-1-8．「All Clear」ボタン

クリップボード、「Current Prompt」フィールド、及び全ての履歴を完全に消去し、リセットします。アンドゥはできません。（通称：Nuke / 物理的更地化ボタン）

#### 5-1-9．重み付け機能（「Weight」チェックボックス＆「Wrap Block」ボタン）

強調したい（あるいは弱めたい）プロンプトを、0.5～1.3まで重み付けできます。一つの単語はもちろん、例えば、先ほどの「oversized + tank top」といった、2単語以上の組み合わせも、「(oversized tank top:1.1)」のようにできます。「Wrap Block」を押すと、その前のカンマまでのブロックを「()」でくくって、重み付けします。もちろん、1つのプロンプトにも有効です。誤操作防止と、AIへの負荷軽減のために、一度「Weight」ボタンをクリックすると、チェックボックスはオフになります。また、プロンプトを追加しないで、もう一度チェックをオンにして「Wrap Block」をクリックすると、重み付けが解除されます。

#### 5-1-10.「Wrap [ ]」ボタン
BREAK構文を使う際、人物ごとの要素丸ごとを”`[ ]`”でくくると、AIの理解がよくなります。（使用するモデルによるかもしれません）そんな時、テキストボックス内の範囲を選択し、このボタンを押すと、そのブロックが”`[ ]`”でくくられます。

#### 5-1-11.「Done!」ボタン
気が済むまで（？）手動での調整が終わったら、「Done!」をクリックすれば、現在のテキストフィードの内容がコピーできます。

#### 5-1-12.「Everyone, Fall in! (Sort Prompt)」ボタン
長大な呪文を構築していると、「要素を入れ忘れる」ことがよくあります。順番がバラバラでも、AIは理解してくれますが、やはり揃っていた方がいい。このボタンを押すと、データベースの項目順に、プロンプトがソートされます。*ただし、`BREAK`構文を使った際は、最後の段落に対してのみ適用される仕様なので、そこはご注意を。*

#### 5-1-11. Send to Favボタン

気に入ったプロンプトを、「Favorite」タブウィンドウのテキストフィールドに転送します。

### 5-2.「Positive」タブ (AIは言われたことしかしない)

基本的に怠け者であるAIを、きちんと働かせるためのポジティブプロンプトを、効率的に管理・運用するための専用画面です。

##### 5-2-1.Your Positive Stock / ポジティブプロンプト貯蔵庫

ウィンドウ左側のリストボックスです。現在登録されている、ポジティブプロンプトが一覧できます。初期状態で、デフォルトがセットされています。これを、足したり引いたり、割ったり掛けたり（？）するわけです。Ctrlキー、あるいはShiftキーを押しながらの、複数選択も可能です。

###### 5-2-1-1.Selected Deleteボタン

ストックデータのうち、選択されたものを削除します。元データも消えます。

###### 5-2-1-2.Delete Allボタン

全てのストックを削除します。元データが全て消えるので、ご注意を。

##### 5-2-2 Applied Negative List / 現在の適用リスト

ウィンドウ右側のリストボックスで、ポジティブプロンプトとして出力したい精鋭ども（？）の場所です。こちらも、複数選択可能でし。でし？

###### 5-2-2-1.Dismiss Allボタン

適用するつもりだった言葉全てを却下（消去）します。元データには影響しません。

##### 5-2-3.▲/▼ ボタン

2つのリストボックス横にあるボタンです。見たまんまで、リスト内の順番をソートします。ただし、エラー防止の観点から、動かせるのは1つずつになります。左側のストックは、元データの順番も変わりますが、右側は影響しません。

##### 5-2-4.→/←ボタン

これも見たままですが、「→」は、ストックから選択した項目を適用リストへ送り、「←」は、適用するつもりだった言葉を戻し（却下）します。元データには影響しません。

##### 5-2-5.Addボタン

選択されたタグを、ウィンドウ下部の「Preview」のテキストボックスへ送ります。複数選択した場合は、「, 」（カンマと半角スペース）で区切られて出力されます。

##### 5-2-6.Set as Defaultボタン

Previewウィンドウの内容を、デフォルトのポジティブプロンプトとして登録します。

##### 5-2-7.Save as Presetボタン

Previewウィンドウの内容を、プリセットとして登録します。

##### 5-2-8.Add New Positive Stock? テキストボックス/ Add Newボタン/Clear Inputボタン

新しいポジティブプロンプトを登録する場合、ここへ入力して、エンターキーか、「Add New」ボタンを押せば、左側のストックリストボックスに追加されます。カンマ区切り、タブ区切り、改行区切りでも、整形して登録できます。Clear Inputは、単純にテキストボックス内をクリアします。

##### 5-2-9.Previewボックス

先に触れましたが、実際のネガティブプロンプトが表示されます。直接入力による編集も可能です。

##### 5-2-10.Send to Cockpitボタン/Clear Positiveボタン

Previewに表示されているポジティブプロンプトを、Cockpitのメインテキストウィンドウへ転送します。この操作は、ショートカットキー`Ctrl＋Shift＋P`でも可能です。Previewウィンドウに何も入っていないと、エラーが出ます。Clear Positiveボタンは、Previewウィンドウをクリアします。

##### 5-2-11.Positive Presetコンボボックス/ Call Presetボタン/ Delete Presetボタン

プルダウンから登録したポジティブプロンプトのプリセットを選択肢、「Call Preset」で呼び出します。Delete Presetでの削除もできますが、(Default)は削除できません。

### 5-3.「Negative」タブ (絶許ブラックリスト)

AIの「暴走」を食い止めるためのネガティブプロンプトを、効率的に管理・運用するための専用画面です。基本的に、ポジティブプロンプトの管理画面とシンメトリーになっていますが、細かくは違います。

##### 5-3-1.Your Negative Stock / ネガティブプロンプト貯蔵庫

ウィンドウ左側のリストボックスです。現在登録されている、ネガティブプロンプトが一覧できます。初期状態で、デフォルトがセットされています。

###### 5-3-1-1.Selected Deleteボタン

ストックデータのうち、選択されたものを削除します。元データも消えます。

###### 5-3-1-2.Delete Allボタン

ポジティブプロンプトの管理画面同様、全てのストックを削除します。元データが全て消えます。

##### 5-3-2 Applied Negative List / 現在の適用リスト

ポジティブプロンプトの管理画面と同様です。

###### 5-3-2-1.Dismiss Allボタン

ポジティブプロンプトの管理画面と（以下同文）

##### 5-3-3.▲/▼ ボタン

ポジティブプロンプトの管理（以下略）

##### 5-4-4.→/←ボタン

ポジティブ（略）

##### 5-4-5.Add or Editボタン

ポジ（略）

##### 5-4-6.Weighting チェックボックス、コンボボックス、Weighten and Addボタン

Cockpit画面同様に、チェックボックスをオンにするとコンボボックスが開くようになり、ネガティブプロンプトを重み付けできます。やはり、0.5～1.3までです。

##### 5-4-7.Add New Negative Stock? テキストボックス/ Add Newボタン/Clear Inputボタン

同じ事を説明させるな！（突然の逆ギレ）

##### 5-4-8.Previewボックス

特に説明はいらないかと。

##### 5-4-9.Copy Negative/Clear Negativeボタン

テキストボックスのネガティブプロンプトを、クリップボードにコピー、あるいは、Previeボックス内を消去します。

### 5-5.「LoRA」 タブ（LoRAの鍛冶場）

品質向上に必須となる &lt;`lora:hash:strength`&gt; 形式のプロンプトと、それに付随するトリガーワードの組み合わせを視覚的かつ直感的に構築・管理します。

（LoRAが一つも登録されていない場合は、安全のためUIがロックされています。「Open LoRA Manage Window」からLoRAを登録してください）

##### 5-5-1. Use LoRA チェックボックス

このタブの機能の有効/無効を切り替えます。

##### 5-5-2. Open LoRA Manage Window ボタン

LoRAのシステムハッシュやトリガーワードを登録・管理する専用ウィンドウ（後述）を開きます。

##### 5-5-3. LoRA Selection & Strength (LoRAの選択と強度)

登録済みのLoRAをプルダウンから選択します。選択すると、推奨の強さ（Strength）と登録されたトリガーワードが自動的に展開されます。

##### 5-5-4. Trigger Words チェックボックス

展開されたトリガーワードから、今回使いたいものだけにチェックを入れます（デフォルトで全てONになる親切設計です）。トリガーワードがないLoRAに関しては、出てきません。

##### 5-5-5. Set! & Cancel ボタン

「Set!」を押すと、選択したLoRAとトリガーワードが右側のリストボックス（カート）に入ります。「Cancel」は選択状態を白紙に戻します。

##### 5-5-6. Your Selected LoRA (Cart)

現在セットされているLoRAのリストです。複数重ね掛けしたい場合は、さらに別のLoRAを選択して「Set!」を押すことでどんどん追加できます。

##### 5-5-7. Wrap LoRA! ボタン & Weight (プロンプトの錬成)

カートに入っている「選択されたLoRAの」トリガーワードを、`&lt;lora:Hash:strength&gt;`,`(Trigger word:1.1)`の形式に整形し、Preview欄に出力します。トリガーワードの重み（Weight）が1.0の場合は、「()」でくくられません。ただし、「同じLoRAの、複数あるトリガーワードに、それぞれ違う重み付けをしたい」場合には、現状の仕様ではできませんので、Previewテキストボックスを手動で編集してください。

- _**Safety Feature:**_ 万が一、大元の管理データから削除された「幽霊LoRA」がカートに残っていた場合、このボタンを押した瞬間に自動検知してカートから除霊（削除）します。

##### 5-5-8. Remove / Forget LoRA ボタン

「Remove」は選択したLoRAをリストから1つ除外します。「Forget」はカートとプレビューを完全にクリアして最初からやり直します。

##### 5-5-9. Send to Cockpit / Send to Fav ボタン

Preview欄に完成したLoRAプロンプトを、Cockpitタブ（またはFavoritesタブ）のテキストフィールドへ転送します。既にCockpitにプロンプトがあっても、先頭に挿入されます。

##### 5-5-10. Preset (プリセットの保存と呼び出し)

「Save as Preset」で現在のLoRAの組み合わせ（リストの中身と重み）に名前を付けて保存できます。「Call Preset」でいつでも一発で呼び出し、「Delete Preset」で削除します。よく使う衣装や画風の組み合わせを保存しておくと劇的に時短になります。

### 5-6. 「LoRA Register & Manage」タブ (LoRAの金庫)

LoRAの実データ（ハッシュ値やトリガーワード）を登録・管理する心臓部です。「LoRA」タブの「Open LoRA Manage Window」からアクセスします。

#### 5-6-1. LoRA Alias / System Hash / Recommended Strength:

- Alias: あなたが識別しやすい任意の名前（例：JK制服、水彩画スタイル）。
- System Hash: SeaArt等で実際に使用されるハッシュ値の文字列。

***Note：ハッシュ値とは？***

SeaArtのサイトの、各LoRAの詳細ページを開いたときのURLの、

https://www.seaart.ai/ja/models/detail/40095be8759dde4285ccf683b24e8852

例えば上記のLoRAの場合、「/detail/」以降の文字列「40095be8759dde4285ccf683b24e8852」が、ハッシュ値です。

ローカル環境のStable Diffusionならば、あなたが持っている`.safetensors`のファイル名です。

- Recommended Strength: そのLoRAの、強さの推奨値
- _**Auto-Sanitize:**_ Hash入力欄に全角文字や不正な記号を入れても、自動的に半角の安全な文字に浄化される鉄壁のガードマンが常駐しています。

#### 5-6-2. LoRA Trigger Word / None チェックボックス:

そのLoRAを発動させるためのトリガーワードを入力し、「Add Trigger」で追加します。複数ある場合はカンマ区切りで入力可能です。

***Note***:トリガーワードは最大10個まで。大文字・小文字を区別しない重複チェックが行われ、プロンプトに有害な記号（! ? @ # ( ) など）はエラーで弾かれます。

トリガーワードが不要なLoRAの場合は「None」にチェックを入れてください。

#### 5-6-3. Register / Cancel / Update LoRA ボタン

入力した内容を登録します。トリガーワードが増えた、あるいはなくなったなど、既存のLoRAを編集（Manage）している最中は、ボタンが「Update LoRA」に変化し、変更がない場合は無駄な上書きを防ぐダーティチェック機能が働きます。

#### 5-6-4. Registered LoRA List & Sort ボタン (▲ / ▼)

登録済みのLoRA一覧です。「▲ (Up)」「▼ (Down)」ボタンを使って、よく使うLoRAを上に固めたり、カテゴリごとに並べ替えたりと、リストを自由に整理できます。ここでの並び順は、メイン画面のプルダウンにも連動します。

#### 5-6-5. Manage / Delete LoRA ボタン

リストから選択したLoRAを編集（Manage）、または完全に削除（Delete）します。

#### 5-6-6. Export / Import LoRA (CSVバックアップ)

- Export: 大切なLoRAコレクションをUTF-8形式のCSVファイルとして書き出します。ファイル名の重複時は自動で枝番が付きます。
- Import: CSVからLoRAデータを読み込みます。既存データへの「追記」が可能で、既に登録されているLoRA（AliasとHashが完全一致するもの）は自動でスキップされる賢いマージ機能を備えています。

### 5-7.「My Favorite」タブ（性癖の金庫室）

苦心の末に（？）作り上げた、大切なプロンプトを管理するためのメニュー類があるタブです。なお、このタブのボタンは全て、「My Favorite」タブがアクティブになっていないと、動作しません。

#### 5-7-1.「Search Fav」フィールド、「Search Fav」ボタン＆「Clear Search」ボタン

フィールドにキーワードを入力して、エンターか「Search Fav」ボタンのクリックで、「My Favorite」シート内を検索できます。検索されるのは「Description」のフィールドのみで、メインシートでのそれ同様、ヒットしたセルが光り、複数ヒットした場合は、ループします。Clear Searchは、単純に、検索ボックスだけを消去します。

#### 5-7-2.「Pull From Cockpit」ボタン

「Cockpit」タブのテキストフィールドに入力されているプロンプトを転写します。

#### 5-7-3.「Send to Cockpit」ボタン

上記とは逆に、Favorite Promptのテキストフィールドの内容を、Cockpitのフィールドへ送ります。お気に入りを、メインシート内のプロンプトを加えることで、更にブラッシュアップするときに。

#### 5-7-4.「Clear Fav」ボタン

「Favorite」タブのフィールドを、全てクリアします。

#### 5-7-5.「Copy Fav」ボタン

「My Favorite」シート上のお気に入りをコピーします。やっぱりね、気に入った呪文は、何度でも使いたいですよね。「コピーされた呪文（プロンプト）は、説明と共にテキストフィールドに表示されます。（説明なしで呪文だけ見て、すぐに思い出せる人も、まずいないでしょう？）

#### 5-7-6.「Tweaked!」ボタン

コピーしたお気に入りプロンプトを、テキストフィールド内にて、さらに手動で調整した後、クリック可能になり、クリックすると、テキストフィールドの内容がコピーされます。

#### 5-7-7.「Add to Fav!」ボタン

フィールド内のプロンプトを、「My Favorite」シートに登録します。「Description」が空欄だと、エラーを吐きます。最大50件まで登録できます。

#### 5-7-8.「Replace Fav」ボタン（コレクションは整理しましょう）

「My Favorite」シートの中で、フォーカスされているセルのプロンプトを、（「Favorite」タブの）テキストフィールドの内容と置換します（セルが空欄ならば、そのまま入力されます）。プロンプトが入っていないセル上でクリックしても、エラーが出ます。例えば、「Favに登録したプロンプトに、新しく要素を追加したら、もっとよくなった！　でも、お気に入りの中に似たものがダブるのは困る！」という場合に使えます。

#### 5-7-9.「Export Fav」ボタン（性癖の遺産と、その保全）

お気に入りリストを、CSVファイルでエクスポートできます。「ナイスな絵が出た！　しかし、あの時のプロンプトって、何だったっけ！？」ということ、よくありますよね？　プロンプトは、ある意味「資産」ですから、それを保全するための機能です。

##### 5-7-9-1.エクスポート（書き出し）の挙動

- **保存形式**： 文字コード UTF-8 のCSV形式で出力されます。
- **ファイル名**：デフォルトで「`MyFavorite_yyyymmdd_HHmm`（日付と時刻）」が設定されます。
- **上書き防止**：保存先に同名のファイルが存在する場合、自動的に「(1)」「(2)」といった枝番が付与されます。既存のバックアップを誤って消去することはありません。

#### 5-7-10「Import Fav」ボタン

CSVファイルから、お気に入りリストをインポートできます。ほら、仲間内で「俺はこんなすごいプロンプトを作ったぜ！」とか、交換したいじゃないですか？　そういうときのための機能です。

##### 5-7-10-1.インポート（取り込み）の挙動

###### 取り込みモードの選択

- 実行時に「既存のデータに追記」するか、「全クリアして新しく取り込む」かを選択できます。

###### データの自動分配

- **1つの行に2つ以上の値がある場合**：1つ目を「Description（説明）」、2つ目を「Prompt（プロンプト）」へ振り分けます。
- **値が1つしかない場合**：データの位置に関わらず、検索の利便性を優先して強制的に「Description」へ格納します。
- **スライド（詰め）機能**： CSV内に形式不良や空行があっても、有効なデータだけを抽出して上方向へ隙間なく詰めてインポートします。
- **日本語のサポート:** 日本語（全角文字）を含む説明も完全にサポートしています。
- **エラーログ:** 制御文字などの不純物により取り込めなかった行がある場合、CSVと同じフォルダに詳細なエラーログ（`ImportLog_時刻.txt`）が生成されます。既存のリストに追加した結果が50件を超える場合、あふれたデータは、エラーログに記録されます。
- ***注意***：インポート元のファイルも、文字コードがUTF-8である必要があります。違っていた場合は、エラーが出て処理されません。
- ***注意2***：仕様上、「データのヘッダを飛ばして4行目から取得を開始する」挙動をします。それより上にあるデータは、取り込めません。

#### 5-7-11.「All Clear Fav」ボタン

「My Favorite」シート内のデータを、全て消去します。元には戻せませんので、くれぐれもご注意を。

### 5-8. 「Gacha!」タブ（AIお任せ・プロンプト錬成）

「自分で呪文を考えるのに疲れたら、Geminiに丸投げすればいいじゃあないか！」（JOJO風に）という、作った作者自身、「よくやったな！？」と思える、ゴイスー（死語）な機能です

つまりは、Google Gemini APIを活用したAIによるプロンプト自動生成機能です 。緻密で複雑な呪文（プロンプト）の構築にうん☆ざりした時の「純粋な息抜き」として、AIの想像力に身を委ねてみるのも、また一興かと。文字通り、「何が出るかな？」の、「ガチャ」です。ちなみに、Geminiにある魔法（誇大表現）をかけてありますので、NSFW絵のプロンプトも、しっかり練成してくれます。ただ、「一発アウト」な地雷ワードもあります。詳しくは、ブック内の「作者的覚え書き(ja)v2」シートをご覧ください。

#### 5-8-1. Google Gemini API Key

AIによる生成機能を利用するには、Google AI Studioにて各自でAPIキーを取得し、入力する必要があります。

- **サポート対象外:** APIキーの取得方法や設定に関する個別サポートは一切行いません。ご自身で解決できる「術師」の方のみご利用ください。作者も、そこまで面倒は見きれません（ぶっちゃけたー！）
- **AIの機嫌:** Google Geminiは非常に「お行儀が良い」AIです。いかにNSFWなプロンプトを出力できる！　とは言え、その内容や（時々サービス精神を強く発揮する）AIの「機嫌（セーフティフィルタ）」次第では、出力が上手く行かなかったり、拒絶されたりすることがあります。それも含めての「ガチャ」としてお楽しみください。
- **APIキーの自動セーブ**:一度Google Geminiによるプロンプト練成が成功すれば、それをフラグにして、入力されたAPIキーは、ローカルマシンのレジストリに記録され、次回からは自動で入力されます。APIキーが変わった場合は、上書きすれば、その時の実行（成功）時に上書きされます。

#### 5-8-2. 自然文入力欄（Please tell me your desire!）

あなたが思い描くシチュエーション(欲望)を、普段使っている言葉で自由に入力してください。

- **バイリンガル対応**: 日本語、英語、あるいはその混在であってもAIが内容を理解し、SeaArt等の生成AIに最適なタグセットへと昇華させます。

#### 5-8-3. 操作ボタン類

- **Feeling Lucky?**: 入力された内容を元に、Geminiへリクエストを送信し、ガチャを回します。エラーでの返答もAPIの使用回数に含められるので、過度な連打はエラーの元＝無駄玉の消費に繋がります。よって、一度押したら、15秒間はグレーアウトします。ちなみにその間、誤操作防止のため、他のタブウィンドウへの移動もできなくなります。バグではなくて、仕様です（いにしえからの常套句）
- **Clear**: 入力欄の内容をクリアします。

#### 5-8-4. 生成結果エリア（How about this?）

AIが錬成したプロンプトが表示されます。

- **Copy This!:** 生成されたプロンプトをクリップボードにコピーします。
- **Send to Cockpit / Send to Fav**: 生成されたプロンプトを「Cockpit」タブまたは「Favorites」タブへ転送します。お気に入りのLoRAを適用したりなど、お好みのままに。
- **Clear Result**: 生成結果を消去します。

#### 5-8-5. 「Trigger Happy?」フレーム（ガチャ残弾数）

本日の残りのガチャ実行回数を表示します。

- **あくまで目安**: 表示される回数は、一般的なGoogleの無料枠でのAPI制限回数と、本ツール内でのカウントに基づいた「目安」です。実際のGoogle APIの制限とは多少の誤差が生じる場合があります。
- **リセット**: カウントは太平洋時間（PT）の0時に合わせてリセットされます。
- **リミット調整（隠し機能）**: APIに課金しているヘビーユーザーが、上限設定を変更したい場合は、カウンターの数字部分をダブルクリックしてください。1日の上限回数を自由に調整できる入力ボックスが表示されます。

#### 5-8-6.「Surprise Me!」チェックボックス / SFW & NSFW & Hardcoreオプションボタン
 - チェックを入れると、3つのオプションボタンが出てきます。どれかを選んでから「Feeling Lucky?」を押すと、データベース内からランダムに抽出された単語を元に、Google Geminiがプロンプトを考えてくれます。

#### Tips：ガチャのコツ

欲望を語るときには、あなたまで「お行儀よく」する必要はないです。変な照れがあると、Geminiがそれを察知して、（小賢しい）安全フィルターを発動させるケースが、デバッグ中にありました。そうです。女性器を指す場合は、「股間」とか、オブラートに包まなくていいんです。素直に「ま○こ」と書きましょう。

## 6\. プロンプトの構築フロー

薄緑色のセルを左から順に選んでいくだけで、一つの完成されたプロンプトになります。

キャラ数(Character Count) → キャラの配置(Character Placement) → 肌の色・属性(Skin & Attributes) → 体型(Body Type) → 髪の毛に関するワイルドカード(Wildcard For Hair) → 髪の長さ(Hair length) → 前髪(Bangs) → 髪の結び目(Tying) → 髪の色(Hair Color) → 体毛(Body Hair) → 職業(Occupation) → 下着(Underwear) → 服装(Outfit) → 服の状態(Outfit State) → ヘッドウェア(Headwear) → 足元周り(Footwear & Legwear) → アクセサリー類(Accessories) → 場所(Location) → 時間帯・周囲の状況(Time & Surroundings) → 体位(Position) → 行為・動作(Action & Movement) → ボンデージ行為(Bondage Action & Movement) → 手段・道具(Means & Props) → 身体の部位(Body Parts) → 行為の状態(Interaction State) → 表情(Expressions) → 体液(Body fluids) → その他アイテム(Misc Items) → アングル(Camera Angle) → 修正その他(Censorship Fixes & Others)

Stable Diffusionのローカル環境で、Dynamic Promptの拡張機能を使用している場合、髪の毛、表情、アングルについては、ワイルドカードを用意しています。ZIPファイル内に同梱していますので、ご活用下さい。

メインシートの冒頭には、各項目のセルへ飛ぶハイパーリンクが設定しており、クリックするとジャンプし、フォーカスの移動したセルは、自動的にウィンドウの左端に寄ります。

## 7,「Sample Prompts」（サンプルプロンプト）シート

作者お気に入りのシチュエーションを収録しています。性癖の開示！（電波）サンプルをコピーすると、「Cockpit」タブウィンドウのテキストボックスにも同時に転送されるので、自分で調節して、オリジナルのプロンプトを作る事も出来ます。ただし、その場合の重み付けなどは、ご自身でお願いします。

## 8.「作者覚え書き(ja)v2 / Author's Notes (Tips and Rants)」シート

プロンプトに関するTipsとか、このマクロを作るに当たっての苦労話とかを書いています。息抜きにどうぞ（？）

## 9.「CONTACT」シート

このREADMEにも書いていますが、作者の連絡先ブログなどを記載しています。

## 10\. 免責事項・連絡先

- 生成結果はAI次第です。当ツールの使用によるいかなる損害についても、作者は一切の責任を負いません。
- 同じく、「Gacha!」の機能についても、NSFWなプロントの、確実な出力を保証するものではありません。
- 検証モデル：[REED_XXX_illustrious_SDXL V14.0](https://civitai.red/models/1717562/reedxxxillustrioussdxl)
- 作者：不二川巴人（ふじかわ ともひと）（「でぇすて」とか、「不二川“でぇすて”巴人」名義で、エロゲーライターをやっていました）
- [連絡先・ブログはこちら。](https://dsblog.biz/)
- リクエストや感想、あるいはバグレポートは、ブログのメールフォームまで。[投げ銭（PayPal）](https://paypal.me/dst0508https:/paypal.me/dst0508)も歓迎です！
- バグレポートももちろんですが、「こんなロケーションが抜けてるぜ！」とか、「この服がないぞ！」というフィードバックは、是非ともお寄せください。自分好みに自由に改変できるとは言え、「未知のシチュエーションを、俺も見たい！」からです！　あなたの意見が、このマクロをよりKENZENにします！
- （おまけ）[SeaArtの個人ページ](https://www.seaart.ai/ja/user/4b23d22e331a382c4adc23a3df4e7077?u_code=XWACJSXI)では、生成したイラストを元に、書き下ろしショートショートを投稿したりしています。よろしければ、そちらもどうぞ。

**さあ！　健やかなる()AIライフを！**

![FLAG_COUNTER](https://s01.flagcounter.com/count2/rmpG/bg_FFFFFF/txt_000000/border_CCCCCC/columns_3/maxflags_12/viewers_0/labels_1/pageviews_1/flags_0/percent_0/)
