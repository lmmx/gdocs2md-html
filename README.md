gdocs2md-html
=============

A simple Google Apps script to convert a properly formatted Google Drive Document to the markdown (.md) format.

## Brief installation

  * Adding this script to your doc (once per doc):
    * Open your Google Drive document (http://drive.google.com)
    * Tools -> Script Editor > Create script for blank project
    * Clear the default Code.gs file and paste the contents of [exportmd.gs](https://raw.githubusercontent.com/lmmx/gdocs2md-html/master/exportmd.gs) into the code editor
    * File -> Save (or `Ctrl`/`⌘` + `S`)
    * When prompted enter new project name for '*Untitled project*', e.g. '*gdocs2md*'

**Note**: now needs advanced Drive API enabling in Script Editor on first use (just [a few clicks from the menu](https://github.com/lmmx/devnotes/wiki/Enabling-advanced-Drive-API)).

* `Resources` menu ⇢ `Advanced Google Services`, enable the Drive API by clicking the switch next to its name
* The link below the API list to the dev console project page for the script you're editing will become clickable
  * At the link, again click the on/off switch on the advanced Drive API setting
* If the script doesn't refresh automatically at this point, hit `Ctrl/Cmd` + `S` to save it, select `onInstall` from the functions dropdown menu at the top of the page and to the left of that click the play icon to run.
* If you go back to your document a `Markdown` menu will be sitting there, with a dropdown to `Export markdown` > `View in browser`

![](https://github.com/lmmx/gdocs2md-html/blob/master/menu.png?raw=true)

> Menu added to Google Docs. Note: customisation/comments features work in progress, see [script](https://github.com/lmmx/gdocs2md-html/blob/master/exportmd.gs)

When you finish writing, the markdown (via the `View in browser` button) can be pasted straight into a suitable publisher.

  * HTML tags are supported in markdown, but feel free to add your own hooks in the source code, or modify the existing ones.
  * Feel free to change the hooks yourself - e.g. the `--- src` ... `---` fences could demarcate `<blockquote>` / `</blockquote>` tags - you're free to edit the script when installing.

Related tips/tricks:

* [devnotes: Custom application to open .gdoc extensions (unix)](https://github.com/lmmx/devnotes/wiki/Custom-application-to-open-.gdoc-extensions)
* [devnotes: Google drive gdoc download conversion workaround](https://github.com/lmmx/devnotes/wiki/Google-drive-gdoc-download-conversion-workaround)

  * Running the script:
    - Select `setupScript` from the script editor's dropdown menu and the play symbol button (on the left of the dropdown) to run it
    - Likewise for the `convertSingleDoc` function, or for batch processing on a folder use the `convertFolder` function.
    - The first run will require you to authorize the app.
    - Converted doc will be saved in the Markdown sub-directory, images below that (`/assets/images`).


## Interpreted formats
  * Text:
    * paragraphs are separated by two newlines
    * text styled as heading 1, 2, 3, etc is converted to Markdown heading: #, ##, ###, etc
    * text formatted with Courier New is backquoted: ``text``
    * links are converted to MD format: `[anchortext](url)`
    * indented paragraphs are rendered as `> ` (`blockquote`) blocks
    * sub- and superscript text is tagged appropriately (`<sup>`/`<sub>`)
  * Lists:
    * Numbered lists are converted correctly, including nested lists
    * bullet lists are converted to "`*`" Markdown format appropriately, including nested lists
  * Images:
    * images are correctly extracted and sent as attachments
  * Blocks:
    * Table of contents is replaced by `[[TOC]]`
    * blocks of text delimited by "--- class whateverclassnameyouwant" and "---" are converted to `<pre class="whateverclassnameyouwant"></pre>` 
    * Source code: 
      * **UPDATED**: blocks of text delimited by "--- source code" or "--- src" and "---" are converted to `<pre></pre>`
      * **NEW**: blocks of text delimited by "--- source pretty" or "--- srcp" and "---" are converted to `<pre class="prettyprint"></pre>`
      * **NEW**: blocks of text delimited by "--- gloss" and "---" are converted to `<pre class="glossary"></pre>`
      * **NEW**: blocks of text delimited by "--- fig-cap" and "---" are converted to `<pre class="fig-cap"></pre>` (figure captions)
       * (feel free to change these for your own needs)
    * Tables:
      * **NEW**: Simple `<table>` processing
  * "--- jsperf `<testID>`" is replaced by an iframe that shows an interactive chart of a JSPerf test. The `<testID>` is the last part of the URL of the Browserscope anchor in your JSPerf test. Something like `"agt1YS1wcm9maWxlcnINCxIEVGVzdBjlm_EQDA"` in the URL `http://www.browserscope.org/user/tests/table/agt1YS1wcm9maWxlcnINCxIEVGVzdBjlm_EQDA`
 


## CONTRIBUTORS

* Renato Mangini - [G+](//google.com/+renatomangini) - [Github](//github.com/mangini)
* Ed Bacher - [G+](//plus.google.com/106923847899206957842) - [Github](//github.com/evbacher)
* Chris Clearfield - [Github](https://github.com/clearf)
* Louis Maddox - [G+](https://plus.google.com/u/0/+LouisMaddox) - [GitHub](https://github.com/lmmx)

## LICENSE

Use this script at your will, on any document you want and for any purpose, commercial or not. 
The MarkDown files generated by this script are not considered derivative work and 
don't require any attribution to the owners of this script. 

If you want to modify and redistribute the script (not the converted documents - those are yours), 
just keep a reference to this repo or to the license info below:

```
Copyright 2013 Google Inc. All Rights Reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
