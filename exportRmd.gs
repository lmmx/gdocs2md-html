/*
Parsing from lmmx/gdocs2Rmd.
Modified by clearf to add files to the google directory structure. 

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

function setupScript() {
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // manual way to do the following:
  // scriptProperties.setProperty("folder_id", "INSERT_FOLDER_ID_HERE");
  // scriptProperties.setProperty("document_id", "INSERT_FILE_ID_HERE");
  
  var doc_id = DocumentApp.getActiveDocument().getId();
  scriptProperties.setProperty("document_id", doc_id);
  var doc_parents = DriveApp.getFileById(doc_id).getParents();
  if ( doc_parents == null ) {
    var folder_id = null; // maybe this is right?
  // TODO: check how to handle docs in root folder - would return null here but at least it's explicit
  } else {
    var folders = doc_parents;
    while (folders.hasNext()) {
      var folder = folders.next();
      var folder_id = folder.getId();
    }
  }
  scriptProperties.setProperty("folder_id", folder_id);
  scriptProperties.setProperty("image_folder_prefix", "/assets/images/");
}

function convertSingleDoc() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var folder_id=scriptProperties.getProperty("folder_id");
  var document_id=scriptProperties.getProperty("document_id");
  var source_folder = DriveApp.getFolderById(folder_id);
  var markdown_folders = source_folder.getFoldersByName("Markdown");

   var markdown_folder; 
  if (markdown_folders.hasNext()) { 
    markdown_folder = markdown_folders.next();
  } else { 
    // Create a markdown folder if it doesn't exist.
    markdown_folder = source_folder.createFolder("Markdown")
  }
  
  convertDocumentToMarkdown(DocumentApp.openById(document_id), markdown_folder);  
}

function convertFolder() {
  var scriptProperties = PropertiesService.getScriptProperties(); 
  var folder_id=scriptProperties.getProperty("folder_id");
  var source_folder = DriveApp.getFolderById(folder_id);
  var markdown_folders = source_folder.getFoldersByName("Markdown");
  
  
  var markdown_folder; 
  if (markdown_folders.hasNext()) { 
    markdown_folder = markdown_folders.next();
  } else { 
    // Create a markdown folder if it doesn't exist.
    markdown_folder = source_folder.createFolder("Markdown");
  }
  
  // Only try to convert google docs files.  
  var gdoc_files = source_folder.getFilesByType("application/vnd.google-apps.document"); 

  // For every file in this directory
  while(gdoc_files.hasNext()) { 
    var gdoc_file = gdoc_files.next()

    var filename = gdoc_file.getName();    
    var Rmd_files = markdown_folder.getFilesByName(filename + ".Rmd");
    var update_file = false
    
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
  var image_prefix=scriptProperties.getProperty("image_folder_prefix");
  var numChildren = document.getActiveSection().getNumChildren();
  var text = "";
  var Rmd_filename = document.getName()+".Rmd";
  var image_foldername = document.getName()+"_images";
  var inSrc = false;
  var inClass = false;
  var globalImageCounter = 0;
  var globalListCounters = {};
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
        text+="<div class=\""+result.className+"\">\n";
      } else if (result.inClass==="end" && inClass) {
        inClass=false;
        text+="</div>\n\n";
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

// Process each child element (not just paragraphs).
function processParagraph(index, element, inSrc, imageCounter, listCounters, image_path) {
  // First, check for things that require no processing.
  if (element.getNumChildren()==0) {
    return null;
  }  
  // Punt on TOC.
  if (element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
    return {"text": "[[TOC]]"};
  }
  
  // Set up for real results.
  var result = {};
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
    var t=element.getChild(i).getType();
    
    if (t === DocumentApp.ElementType.TABLE_ROW) {
      // do nothing: already handled TABLE_ROW
    } else if (t === DocumentApp.ElementType.TEXT) {
      var txt=element.getChild(i);
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
      textElements.push('![image alt text](' + image_path + '/' + name + ')');
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

    // replace Unicode quotation marks
    pOut = pOut.replace('\u201d', '"').replace('\u201c', '"');
 
    result.text = prefix+pOut;
  }
  
  return result;
}

// Add correct prefix to list items.
function findPrefix(inSrc, element, listCounters) {
  var prefix="";
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
  
  var attrs=txt.getTextAttributeIndices();
  var lastOff=pOut.length;
  // We will run through this loop multiple times for the things we care about.
  // Font
  // URL
  // Then for bold
  // Then for italiac.
  
  // FONTs
  var lastOff=pOut.length;
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
