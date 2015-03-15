/*
Parsing from mangini/gdocs2md.
Modified by clearf to add files to the google directory structure. 
Modified by lmmx to write Markdown, going back to HTML-incorporation.

Usage: 
  NB: don't use on top-level doc (in root Drive folder) See comment in setupScript function.
  Adding this script to your doc: 
    - Tools > Script Manager > New
    - Select "Blank Project", then paste this code in and save.
  Running the script:
    - Tools > Script Manager
    - Select "convertDocumentToMarkdown" function.
    - Click Run button.
    - Converted doc will be added to a "Markdown" folder in the source document's directories. 
    - Images will be added to a subfolder of the "Markdown" folder. 
*/

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  // Add a menu with some items, some separators, and a sub-menu.
  setupScript();
// In future:
//  DocumentApp.getUi().createAddonMenu();
  DocumentApp.getUi().createMenu('Markdown')
      .addItem('Export \u2192 markdown', 'convertSingleDoc')
      .addItem('Export folder \u2192 markdown', 'convertFolder')
      .addItem('Customise markdown conversion', 'changeDefaults')
      .addSeparator()
      .addSubMenu(DocumentApp.getUi().createMenu('Toggle comment visibility')
                 .addItem('Image source URLs', 'toggleImageSourceStatus')
                 .addItem('All comments', 'toggleCommentStatus'))
      .addItem("Add comment", 'addCommentDummy')
      .addToUi();
}

function changeDefaults() {
  var ui = DocumentApp.getUi();
  var default_settings = '{ use your imagination... }';
  var greeting = ui.alert('This should be set up to display defaults from variables passed to getDocComments etc., e.g. something like:\n\nDefault settings are:'
                    + '\ncomments - not checking deleted comments.\nDocument - this document (alternatively specify a document ID).'
                    + '\n\nClick OK to edit these, or cancel.',
                    ui.ButtonSet.OK_CANCEL);
  ui.alert("There's not really need for this yet, so this won't proceed, regardless of what you just pressed.");
  return;
  
  // Future:
  if (greeting == ui.Button.CANCEL) {
    ui.alert("Alright, never mind!");
    return;
  }
  // otherwise user clicked OK
  // user clicked OK, to proceed with editing these defaults. Ask case by case whether to edit
  
  var response = ui.prompt('What is x (default y)?', ui.ButtonSet.YES_NO_CANCEL);
 
  // Example code from docs at https://developers.google.com/apps-script/reference/base/button-set
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    Logger.log('The user\'s name is %s.', response.getResponseText());
  } else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function setupScript() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("user_email", Drive.About.get().user.emailAddress);
  
  // manual way to do the following:
  // scriptProperties.setProperty("folder_id", "INSERT_FOLDER_ID_HERE");
  // scriptProperties.setProperty("document_id", "INSERT_FILE_ID_HERE");
  
  var doc_id = DocumentApp.getActiveDocument().getId();
  scriptProperties.setProperty("document_id", doc_id);
  var doc_parents = DriveApp.getFileById(doc_id).getParents();
  var folders = doc_parents;
  while (folders.hasNext()) {
    var folder = folders.next();
    var folder_id = folder.getId();
  }
  scriptProperties.setProperty("folder_id", folder_id);
  scriptProperties.setProperty("image_folder_prefix", "/images/");
}

function addCommentDummy() {
  // Dummy function to be switched during development for addComment  
  DocumentApp.getUi()
    .alert('Cancelling comment entry',
           "There's not currently a readable anchor for Google Docs - you need to write your own!"
           
           + "\n\nThe infrastructure for using such an anchoring schema is sketched out in"
           + " the exportmd.gs script's addComment function, for an anchor defined in anchor_props"
           
           + "\n\nSee github.com/lmmx/devnotes/wiki/Custom-Google-Docs-comment-anchoring-schema",
           DocumentApp.getUi().ButtonSet.OK
          );
  return;
}

function addComment() {
  
  var doc_id = PropertiesService.getScriptProperties().getProperty('document_id');
  var user_email = PropertiesService.getScriptProperties().getProperty('email');
/*  Drive.Comments.insert({content: "hello world",
                         context: {
                           type: 'text/html',
                           value: 'hinges'
                         }
                        }, document_id); */
  var revision_list = Drive.Revisions.list(doc_id).items;
  var recent_revision_id = revision_list[revision_list.length - 1].id;
  var anchor_props = {
    revision_id: recent_revision_id,
    starting_offset: '',
    offset_length: '',
    total_chars: ''
  }
  insertComment(doc_id, 'hinges', 'Hello world!', my_email, anchor_props);
}
  
function insertComment(fileId, selected_text, content, user_email, anchor_props) {
  
  /*
  anchor_props is an object with 4 properties:
    - revision_id,
    - starting_offset,
    - offset_length,
    - total_chars
  */
  
  var context = Drive.newCommentContext();
    context.value = selected_text;
    context.type = 'text/html';
  var comment = Drive.newComment();
    comment.kind = 'drive#comment';
    var author = Drive.newUser();
      author.kind = 'drive#user';
      author.displayName = user_email;
      author.isAuthenticatedUser = true;
    comment.author = author;
    comment.content = type; 
    comment.context = context;
    comment.status = 'open';
    comment.anchor = "{'r':"
                     + anchor_props.revision_id
                     + ",'a':[{'txt':{'o':"
                     + anchor_props.starting_offset
                     + ",'l':"
                     + anchor_props.offset_length
                     + ",'ml':"
                     + anchor_props.total_chars
                     + "}}]}";
    comment.fileId = fileId;
  Drive.Comments.insert(comment, fileId);
}

function getDocComments(comment_list_args) {
  if (typeof(comment_list_args) == 'undefined') {
    var comment_list_args = new Object();
  }
  
  var possible_args = ['images', 'include_deleted'];
  for (var i in possible_args) {
    var possible_arg = possible_args[i];
    if (comment_list_args.propertyIsEnumerable(possible_arg)) {
      eval(possible_arg + " = " + comment_list_args[possible_arg]);
    } else {
      eval(possible_arg + " = " + false);
    }
  }
  
  /*
  Looks bad but more sensible than repeatedly checking if arg undefined.
  
  Sets every variable named in the possible_args array to false if
  it wasn't passed into the comment_list_args object.
  */
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var document_id = scriptProperties.getProperty("document_id");
  var comments_list = Drive.Comments.list(document_id,
                                          {includeDeleted: include_deleted,
                                           maxResults: 100 }); // 0 to 100, default 20
  // See https://developers.google.com/drive/v2/reference/comments/list for all options
  var comment_array = [];
  for (var i = 0; i < comments_list.items.length; i++) {
    var comment_text = comments_list.items[i].content;
 /*
    images is a generic parameter passed in as a switch to
    return image URL-containing comments only.
    
    If the parameter is provided, it's no longer undefined.
 */
    if (images) {
      if (/(https?:\/\/.+?\.(png|gif|jpe?g))/.test(comment_text)) {
        comment_array.push(RegExp.$1);
      } // otherwise there's no image URL here, skip it
    } else { 
      comment_array.push(comment_text);
    }
  }
  return comment_array;
}

function toggleCommentStatus(comment_switches){
  // getDocComments(content_switches);
}

function toggleImageSourceStatus(){
  toggleCommentStatus({images: true});
}

function getImageComments() {
  // for testing/maybe easy shorthand
  return getDocComments({images: true});
}

function flipResolved() {
  // Flip the status of resolved comments to open, and open comments to resolved (respectful = true)
  // I.e. make resolved URL-containing comments visible, without losing track of normal comments' status
  
  // To force all comments' statuses to switch between resolved and open en masse set respectful to false
  
  var switch_settings = new Object();
    switch_settings.respectful = true;
    switch_settings.images_only = false; // If true, only switch status of comments with an image URL
    switch_settings.switch_deleted_comments = false; // If true, also switch status of deleted comments
  
  var comments_list = getDocComments(
    { images: switch_settings.images_only,
      include_deleted: switch_settings.switch_deleted_comments });
  
  // Note: these parameters are unnecessary if both false (in their absence assumed false)
  //       but included for ease of later reuse
    
  if (switch_settings.respectful) {
    // flip between
  } else {
    // flip all based on status of first in list
  }
}

function convertSingleDoc() {
  var scriptProperties = PropertiesService.getScriptProperties();
  // renew comments list on every export
  var doc_comments = getDocComments();
  var image_urls = getDocComments({images: true}); // NB assumed false - any value will do
  scriptProperties.setProperty("comments", doc_comments);
  scriptProperties.setProperty("image_srcs", image_urls);
  var folder_id = scriptProperties.getProperty("folder_id");
  var document_id = scriptProperties.getProperty("document_id");
  var source_folder = DriveApp.getFolderById(folder_id);
  var markdown_folders = source_folder.getFoldersByName("Markdown");

  var markdown_folder; 
  if (markdown_folders.hasNext()) { 
    markdown_folder = markdown_folders.next();
  } else { 
    // Create a Markdown folder if it doesn't exist.
    markdown_folder = source_folder.createFolder("Markdown")
  }
  
  convertDocumentToMarkdown(DocumentApp.openById(document_id), markdown_folder);  
}

function convertFolder() {
  var scriptProperties = PropertiesService.getScriptProperties(); 
  var folder_id = scriptProperties.getProperty("folder_id");
  var source_folder = DriveApp.getFolderById(folder_id);
  var markdown_folders = source_folder.getFoldersByName("Markdown");
  
  
  var markdown_folder; 
  if (markdown_folders.hasNext()) { 
    markdown_folder = markdown_folders.next();
  } else { 
    // Create a Markdown folder if it doesn't exist.
    markdown_folder = source_folder.createFolder("Markdown");
  }
  
  // Only try to convert google docs files.  
  var gdoc_files = source_folder.getFilesByType("application/vnd.google-apps.document"); 

  // For every file in this directory
  while(gdoc_files.hasNext()) { 
    var gdoc_file = gdoc_files.next()

    var filename = gdoc_file.getName();    
    var Rmd_files = markdown_folder.getFilesByName(filename + ".Rmd");
    var update_file = false;
    
    if (Rmd_files.hasNext()) {
      var Rmd_file = Rmd_files.next();
      
      if (Rmd_files.hasNext()){ // There are multiple markdown files; delete and rerun
        update_file = true;
      } else if (Rmd_file.getLastUpdated() < gdoc_file.getLastUpdated()) { 
        update_file = true; 
      }
    } else {
      // There is no folder and the conversion needs to be rerun
      update_file = true;
    }  
    
    if (update_file) { 
      convertDocumentToMarkdown(DocumentApp.openById(gdoc_file.getId()), markdown_folder);
    }
  }
}

function convertDocumentToMarkdown(document, destination_folder) {
  var scriptProperties = PropertiesService.getScriptProperties(); 
  var image_prefix = scriptProperties.getProperty("image_folder_prefix");
  var numChildren = document.getActiveSection().getNumChildren();
  var text = "";
  var Rmd_filename = document.getName()+".Rmd";
  var image_foldername = document.getName()+"_images";
  var inSrc = false;
  var inClass = false;
  var globalImageCounter = 0;
  var globalListCounters = new Object();
  // edbacher: added a variable for indent in src <pre> block. Let style sheet do margin.
  var srcIndent = "";
  
  var postHasImages = false; 
  
  var files = [];
  
  // Walk through all the child elements of the doc.
  for (var i = 0; i < numChildren; i++) {
    var child = document.getActiveSection().getChild(i);
    var result = processParagraph(i, child, inSrc, globalImageCounter, globalListCounters, image_prefix + image_foldername);
    globalImageCounter += (result && result.images) ? result.images.length : 0;
    if (result!==null) {
      if (result.sourcePretty==="start" && !inSrc) {
        inSrc=true;
        text+="<pre class=\"prettyprint\">\n";
      } else if (result.sourcePretty==="end" && inSrc) {
        inSrc=false;
        text+="</pre>\n\n";
      } else if (result.source==="start" && !inSrc) {
        inSrc=true;
        text+="<pre>\n";
      } else if (result.source==="end" && inSrc) {
        inSrc=false;
        text+="</pre>\n\n";
      } else if (result.inClass==="start" && !inClass) {
        inClass=true;
        text+="<pre class=\""+result.className+"\">\n";
      } else if (result.inClass==="end" && inClass) {
        inClass=false;
        text+="</pre>\n\n";
      } else if (inClass) {
        text+=result.text+"\n\n";
      } else if (inSrc) {
        text+=(srcIndent+escapeHTML(result.text)+"\n");
      } else if (result.text && result.text.length>0) {
        text+=result.text+"\n\n";
      }
      
      if (result.images && result.images.length>0) {
        for (var j=0; j<result.images.length; j++) {
          files.push( { "blob": result.images[j].blob } );
          postHasImages = true; 
        }
      }
    } else if (inSrc) { // support empty lines inside source code
      text+='\n';
    }
      
  }
  files.push({"fileName": Rmd_filename, "mimeType": "text/plain", "content": text});
    
  
  // Cleanup any old folders and files in our destination directory with an identical name
  var old_folders = destination_folder.getFoldersByName(image_foldername)
  while (old_folders.hasNext()) {
    var old_folder = old_folders.next();
    old_folder.setTrashed(true)
  }  
  
  // Remove any previously converted markdown files.
  var old_files = destination_folder.getFilesByName(Rmd_filename)
  while (old_files.hasNext()) {
    var old_file = old_files.next();
    old_file.setTrashed(true)
  }  
  
  // Create a subfolder for images if they exist
  var image_folder; 
  if (postHasImages) { 
    image_folder = DriveApp.createFolder(image_foldername);
    DriveApp.removeFolder(image_folder); // Confusing convention; this just removes the folder from the google drive root.
    destination_folder.addFolder(image_folder)
  }
  
  for (var i = 0; i < files.length; i++) { 
    var saved_file; 
    if (files[i].blob) {
      saved_file = DriveApp.createFile(files[i].blob)
      // The images go into a subfolder matching the post title
      image_folder.addFile(saved_file)
    } else { 
      // The markdown files all go in the "Markdown" directory
      saved_file = DriveApp.createFile(files[i]["fileName"], files[i]["content"], files[i]["mimeType"])  
      destination_folder.addFile(saved_file)
    }
   DriveApp.removeFile(saved_file) // Removes from google drive root.
  }
  
}

function escapeHTML(text) {
  return text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function standardQMarks(text) {
  return text.replace(/\u2018|\u8216|\u2019|\u8217/g,"'").replace(/\u201c|\u8220|\u201d|\u8221/g, '"')
}

// Process each child element (not just paragraphs).
function processParagraph(index, element, inSrc, imageCounter, listCounters, image_path) {
  // First, check for things that require no processing.
  if (element.getNumChildren()==0) {
    return null;
  }  
  // Skip on TOC.
  if (element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
    return {"text": "[[TOC]]"};
  }
  
  // Set up for real results.
  var result = new Object();
  var pOut = "";
  var textElements = [];
  var imagePrefix = "image_";
  
  // Handle Table elements. Pretty simple-minded now, but works for simple tables.
  // Note that Markdown does not process within block-level HTML, so it probably 
  // doesn't make sense to add markup within tables.
  if (element.getType() === DocumentApp.ElementType.TABLE) {
    textElements.push("<table>\n");
    var nCols = element.getChild(0).getNumCells();
    for (var i = 0; i < element.getNumChildren(); i++) {
      textElements.push("  <tr>\n");
      // process this row
      for (var j = 0; j < nCols; j++) {
        textElements.push("    <td>" + element.getChild(i).getChild(j).getText() + "</td>\n");
      }
      textElements.push("  </tr>\n");
    }
    textElements.push("</table>\n");
  }
  
  // Process various types (ElementType).
  for (var i = 0; i < element.getNumChildren(); i++) {
    var t = element.getChild(i).getType();
    
    if (t === DocumentApp.ElementType.TABLE_ROW) {
      // do nothing: already handled TABLE_ROW
    } else if (t === DocumentApp.ElementType.TEXT) {
      var txt  = element.getChild(i);
      pOut += txt.getText();
      textElements.push(txt);
    } else if (t === DocumentApp.ElementType.INLINE_IMAGE) {
      result.images = result.images || [];
      var blob = element.getChild(i).getBlob()
      var contentType = blob.getContentType();
      var extension = "";
      if (/\/png$/.test(contentType)) {
        extension = ".png";
      } else if (/\/gif$/.test(contentType)) {
        extension = ".gif";
      } else if (/\/jpe?g$/.test(contentType)) {
        extension = ".jpg";
      } else {
        throw "Unsupported image type: "+contentType;
      }
     
      var name = imagePrefix + imageCounter + extension;
      blob.setName(name);
      
      imageCounter++;
      textElements.push('![](' + image_path + '/' + name + ')');
      //result.images.push( {
      //  "bytes": blob.getBytes(), 
      //  "type": contentType, 
      //  "name": name});
      
      result.images.push({ "blob" : blob } )
      
    } else if (t === DocumentApp.ElementType.PAGE_BREAK) {
      // ignore
    } else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {
      textElements.push('* * *\n');
    } else if (t === DocumentApp.ElementType.FOOTNOTE) {
      textElements.push(' ('+element.getChild(i).getFootnoteContents().getText()+')');
    } else {
      throw "Paragraph "+index+" of type "+element.getType()+" has an unsupported child: "
      +t+" "+(element.getChild(i)["getText"] ? element.getChild(i).getText():'')+" index="+index;
    }
  }

  if (textElements.length==0) {
    // Isn't result empty now?
    return result;
  }
  
  // evb: Add source pretty too. (And abbreviations: src and srcp.)
  // process source code block:
  if (/^\s*---\s+srcp\s*$/.test(pOut) || /^\s*---\s+source pretty\s*$/.test(pOut)) {
    result.sourcePretty = "start";
  } else if (/^\s*---\s+src\s*$/.test(pOut) || /^\s*---\s+source code\s*$/.test(pOut)) {
    result.source = "start";
  } else if (/^\s*---\s+class\s+([^ ]+)\s*$/.test(pOut)) {
    result.inClass = "start";
    result.className = RegExp.$1;
  } else if (/^\s*---\s*$/.test(pOut)) {
    result.source = "end";
    result.sourcePretty = "end";
    result.inClass = "end";
  } else if (/^\s*---\s+jsperf\s*([^ ]+)\s*$/.test(pOut)) {
    result.text = '<iframe style="width: 100%; height: 340px; overflow: hidden; border: 0;" '+
                  'src="http://www.html5rocks.com/static/jsperfview/embed.html?id='+RegExp.$1+
                  '"></iframe>';
  } else {

    prefix = findPrefix(inSrc, element, listCounters);
  
    var pOut = "";
    for (var i=0; i<textElements.length; i++) {
      pOut += processTextElement(inSrc, textElements[i]);
    }

    // replace Unicode quotation marks (double and single)
    pOut = standardQMarks(pOut);
 
    result.text = prefix+pOut;
  }
  
  return result;
}

// Add correct prefix to list items.
function findPrefix(inSrc, element, listCounters) {
  var prefix = "";
  if (!inSrc) {
    if (element.getType()===DocumentApp.ElementType.PARAGRAPH) {
      var paragraphObj = element;
      switch (paragraphObj.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
        case DocumentApp.ParagraphHeading.HEADING6: prefix+="#";
        case DocumentApp.ParagraphHeading.HEADING5: prefix+="#";
        case DocumentApp.ParagraphHeading.HEADING4: prefix+="#";
        case DocumentApp.ParagraphHeading.HEADING3: prefix+="#";
        case DocumentApp.ParagraphHeading.HEADING2: prefix+="#";
        case DocumentApp.ParagraphHeading.HEADING1: prefix+="# ";
        default:
      }
    } else if (element.getType()===DocumentApp.ElementType.LIST_ITEM) {
      var listItem = element;
      var nesting = listItem.getNestingLevel()
      for (var i=0; i<nesting; i++) {
        prefix += "    ";
      }
      var gt = listItem.getGlyphType();
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix += "* ";
      } else {
        // Ordered list (<ol>):
        var key = listItem.getListId() + '.' + listItem.getNestingLevel();
        var counter = listCounters[key] || 0;
        counter++;
        listCounters[key] = counter;
        prefix += counter+". ";
      }
    }
  }
  return prefix;
}

function processTextElement(inSrc, txt) {
  if (typeof(txt) === 'string') {
    return txt;
  }
  
  var pOut = txt.getText();
  if (! txt.getTextAttributeIndices) {
    return pOut;
  }
  
  Logger.log("Initial String: " + pOut)
 
  // CRC introducing reformatted_txt to let us apply rational formatting that we can actually parse
  var reformatted_txt = txt.copy();
  reformatted_txt.deleteText(0,pOut.length-1); 
  reformatted_txt = reformatted_txt.setText(pOut); 
  
  var attrs = txt.getTextAttributeIndices();
  var lastOff = pOut.length;
  // We will run through this loop multiple times for the things we care about.
  // Font
  // URL
  // Then for bold
  // Then for italiac.
  
  // FONTs
  var lastOff = pOut.length;
  for (var i=attrs.length-1; i>=0; i--) {
    var off=attrs[i];
    var font=txt.getFontFamily(off)
     if (font) {
       while (i>=1 && txt.getFontFamily(attrs[i-1])==font) {
          // detect fonts that are in multiple pieces because of errors on formatting:
          i-=1;
          off=attrs[i];
       }
       reformatted_txt.setFontFamily(off, lastOff-1, font); 
     }
    lastOff=off;  
  }
  
  // URL
  // XXX TODO actually convert to URL text here. 
  var lastOff=pOut.length;
  for (var i=attrs.length-1; i>=0; i--) {
    var off=attrs[i];
    var url=txt.getLinkUrl(off);
     if (url) {
       while (i>=1 && txt.getLinkUrl(attrs[i-1]) == url) {
          // detect urls that are in multiple pieces because of errors on formatting:
          i-=1;
          off=attrs[i];
       }
     reformatted_txt.setLinkUrl(off, lastOff-1, url); 
     }
    lastOff=off;  
  }  
  
   // bold
  var lastOff=pOut.length;
  for (var i=attrs.length-1; i>=0; i--) {
    var off=attrs[i];
    var bold=txt.isBold(off);
     if (bold) {
       while (i>=1 && txt.isBold(attrs[i-1])) {
          i-=1;
          off=attrs[i];
       }
     reformatted_txt.setBold(off, lastOff-1, bold); 
     }
    lastOff=off;  
  }
  
  // italics
  var lastOff=pOut.length;
  for (var i=attrs.length-1; i>=0; i--) {
    var off=attrs[i];
    var italic=txt.isItalic(off);
     if (italic) {
       while (i>=1 && txt.isItalic(attrs[i-1])) {
          i-=1;
          off=attrs[i];
       }
     reformatted_txt.setItalic(off, lastOff-1, italic); 
     }
    lastOff=off;  
  }
  
  
  var mOut=""; // Modified out string
  var harmonized_attrs = reformatted_txt.getTextAttributeIndices();
  reformatted_txt.getTextAttributeIndices();
  pOut = reformatted_txt.getText(); 
  
  
  // Markdown is farily picky about how it will let you intersperse spaces around words and strong/italics chars. This regex (hopefully) clears this up
  // Match any number of \*, followed by spaces/workd boundaries against anything that is not the \*, followed by boundaries, spaces and * again. 
  // Test case at http://jsfiddle.net/ovqLv0s9/2/

  var reAlignStars = /(\*+)(\s*\b)([^\*]+)(\b\s*)(\*+)/g;
  
  var lastOff=pOut.length;
  for (var i=harmonized_attrs.length-1; i>=0; i--) {
    var off=harmonized_attrs[i];
    
    var raw_text = pOut.substring(off, lastOff) 
 
    var d1 = ""
    var d2 = ""; 
    
    var end_font; 
    
    var mark_bold = false; 
    var mark_italic =  false; 
    var mark_code = false;
        
    // The end of the text block is a special case. 
    if (lastOff == pOut.length) {  
      end_font = reformatted_txt.getFontFamily(lastOff - 1)
      if (end_font) {
        if (!inSrc && end_font===end_font.COURIER_NEW) {
          mark_code = true;  
        }
      }
      if (reformatted_txt.isBold(lastOff -1)) { 
        mark_bold = true;
      }
      if (reformatted_txt.isItalic(lastOff - 1)) {
        // edbacher: changed this to handle bold italic properly.
        mark_italic = true; 
      }
    } else { 
      end_font = reformatted_txt.getFontFamily(lastOff -1 )
      if (end_font) {
        if (!inSrc && end_font===end_font.COURIER_NEW && reformatted_txt.getFontFamily(lastOff) != end_font) {
          mark_code=true;
        }
      }
      if (reformatted_txt.isBold(lastOff - 1) && !reformatted_txt.isBold(lastOff) ) { 
        mark_bold=true;
      }
      if (reformatted_txt.isItalic(lastOff - 1) && !reformatted_txt.isItalic(lastOff)) {
        mark_italic=true; 
      }
    }
    
    if (mark_code) { 
      d2 = '`';  
    }
    if (mark_bold) { 
      d2 = "**" + d2; 
    }
    if (mark_italic) {
      d2 = "*" + d2; 
    }
    
    mark_bold = mark_italic =  mark_code = false; 
    
    var font=reformatted_txt.getFontFamily(off);
   if (off == 0) {   
      if (font) {
        if (!inSrc && font===font.COURIER_NEW) {
          mark_code = true;  
        }
      }
      if (reformatted_txt.isBold(off)) { 
        mark_bold = true;
      }
      if (reformatted_txt.isItalic(off)) {
        mark_italic = true; 
      }
    } else { 
      if (font) {
        if (!inSrc && font===font.COURIER_NEW && reformatted_txt.getFontFamily(off - 1) != font) {
          mark_code=true;
        }
      }
      if (reformatted_txt.isBold(off) && !reformatted_txt.isBold(off -1) ) { 
        mark_bold=true;
      }
      if (reformatted_txt.isItalic(off) && !reformatted_txt.isItalic(off - 1)) {
        mark_italic=true; 
      }
    }
        
    if (mark_code) {
        d1 = '`'; 
    }    
    
    if (mark_bold) {
      d1 = d1 + "**"; 
    }
    
    if (mark_italic) {
      d1 = d1 + "*"; 
    }
    
    var url=reformatted_txt.getLinkUrl(off);
    if (url) {
      mOut = d1 + '['+ raw_text +']('+url+')' + d2 + mOut;
    } else { 
      var new_text = d1 + raw_text + d2;     
      new_text = new_text.replace(reAlignStars, "$2$1$3$5$4");
      mOut =  new_text + mOut;
    }
      
    lastOff=off;  
    Logger.log("Modified String: " + mOut)
  }
  
  mOut = pOut.substring(0, off) + mOut;    
  return mOut;
}
