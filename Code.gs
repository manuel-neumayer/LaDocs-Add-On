function onOpen() {
  let ui = DocumentApp.getUi();
  let latexmenue = ui.createMenu("LaTex")
  latexmenue.addItem("Convert", "convert").addToUi()
  latexmenue.addItem("Reset to LaTex", "reset").addToUi()
  latexmenue.addItem("Open Sidebar", "showSidebar").addToUi()
  ui.createAddonMenu().addItem("Convert LaTex", "convert").addItem("Reset LaTex", "reset").addToUi()
  PropertiesService.getDocumentProperties().setProperty("bigsize", "25")
  PropertiesService.getDocumentProperties().setProperty("sidebartext", "")
}

function onTheFly() {
  let document = DocumentApp.getActiveDocument()
  let body = document.getBody()
  if (body.findText("§§") !== null) {
    let paragraphs = body.getParagraphs()
    for (si=0; si<paragraphs.length;si++) {
      if (paragraphs[si].getType() == DocumentApp.ElementType.PARAGRAPH && paragraphs[si].findText("§§") !== null) {
        for (sj=0; sj<paragraphs[si].getNumChildren(); sj++) {
          var child = paragraphs[si].getChild(sj)
          if (child.getType() == DocumentApp.ElementType.TEXT && child.asText().findText("§") !== null && child.asText().findText("§§", child.asText().findText("§")) !== null) {
            searchAndInsertFormula(child, paragraphs[si], "§§")
          }
        }
      }
    }
  }
  let listitems = body.getListItems()
  for (si=0; si<listitems.length;si++) {
  if (listitems[si].getType() == DocumentApp.ElementType.LIST_ITEM && listitems[si].findText("§§") !== null) {
      for (sj=0; sj<listitems[si].getNumChildren(); sj++) {
        var child = listitems[si].getChild(sj)
        if (child.getType() == DocumentApp.ElementType.TEXT && child.asText().findText("§") !== null && child.asText().findText("§§", child.asText().findText("§")) !== null) {
          searchAndInsertFormula(child, listitems[si], "§§")
        }
      }
    }
  }
  fixSidebarImage()
}

function convert() {
  onTheFly()
  let document = DocumentApp.getActiveDocument()
  let body = document.getBody()
  if (body.findText("§") !== null) {
    let paragraphs = body.getParagraphs()
    for (si=0; si<paragraphs.length;si++) {
      if (paragraphs[si].getType() == DocumentApp.ElementType.PARAGRAPH && paragraphs[si].findText("§") !== null) {
        for (sj=0; sj<paragraphs[si].getNumChildren(); sj++) {
          var child = paragraphs[si].getChild(sj)
          if (child.getType() == DocumentApp.ElementType.TEXT && child.asText().findText("§") !== null && child.asText().findText("§", child.asText().findText("§")) !== null) {
            if (child.asText().findText("§§", child.asText().findText("§")) == null || child.asText().findText("§", child.asText().findText("§")).getStartOffset() !== child.asText().findText("§§", child.asText().findText("§")).getStartOffset()) {
              searchAndInsertFormula(child, paragraphs[si])
            }
          }
        }
      }
    }
  }
  let listitems = body.getListItems()
  for (si=0; si<listitems.length;si++) {
  if (listitems[si].getType() == DocumentApp.ElementType.LIST_ITEM && listitems[si].findText("§") !== null) {
      for (sj=0; sj<listitems[si].getNumChildren(); sj++) {
        var child = listitems[si].getChild(sj)
        if (child.getType() == DocumentApp.ElementType.TEXT && child.asText().findText("§") !== null && child.asText().findText("§", child.asText().findText("§")) !== null) {
          if (child.asText().findText("§§", child.asText().findText("§")) == null || child.asText().findText("§", child.asText().findText("§")).getStartOffset() !== child.asText().findText("§§", child.asText().findText("§")).getStartOffset()) {
            searchAndInsertFormula(child, paragraphs[si])
          }
        }
      }
    }
  }
}

function searchAndInsertFormula(child, textparent, searchtype) {
  if (searchtype == undefined) {
    searchtype = "§"
  }
  child = child.asText()
  //Let's see where exactly the formula is hiding:
  let startindex = child.findText("§")
  let counter = 0
  while (child.findText("§", startindex) !== null && child.findText("§", startindex).getStartOffset() < child.findText(searchtype, startindex).getStartOffset() && counter < 1000) {
    startindex = child.findText("§", startindex)
    counter++
  }
  let endindex = child.findText(searchtype, startindex)
  //The idea is now to let the text element end at the first dollar sign, insert the formula image as a new child to the paragraph, and insert a new text element child with the rest of the text-string
  let prevtext = child.getText().slice(0, startindex.getEndOffsetInclusive())
  let formula = child.getText().slice(startindex.getEndOffsetInclusive()+1, endindex.getEndOffsetInclusive())
  if (searchtype == "§§") {
    formula = formula.slice(0, formula.length - 1)
  }
  console.log(formula)
  let imagefetch = UrlFetchApp.fetch(encodeURI("https://latex.codecogs.com/png.download?" + formula), {"muteHttpExceptions": true})
  console.log(encodeURI("https://latex.codecogs.com/png.download?" + formula))
  if (imagefetch.getResponseCode() !== 200) {
    return
  }
  let image = imagefetch.getBlob()
  child.deleteText(0, endindex.getEndOffsetInclusive())
  let picture = textparent.insertInlineImage(sj, image)
  picture.setAltDescription("§" + formula + "§")
  picture.setLinkUrl(encodeURI("www.codecogs.com/eqnedit.php?latex=" + formula))
  let size = 11 * 1.25
  if (textparent.getNumChildren() == 2 && prevtext == "" && child.getText() == "") {
    console.log(PropertiesService.getDocumentProperties())
    size = parseFloat(PropertiesService.getDocumentProperties().getProperty("bigsize")) * 1.25
    textparent.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  } else if (child.getAttributes().FONT_SIZE !== null) {
    size = child.getFontSize() * 1.25
  }
  picture.setWidth(picture.getWidth() * (size / picture.getHeight()))
  picture.setHeight(size)
  if (prevtext !== "") {
    textparent.insertText(sj, prevtext)
  }
}

function reset() {
  let selection = DocumentApp.getActiveDocument().getSelection()
  if (selection !== null) {
    resetSelection(selection)
  } else {
    resetAll()
  }
}

function resetSelection(selection) {
  var elements = selection.getRangeElements();
  for (var i = 0; i < elements.length; i++) {
    let element = elements[i].getElement()
    if (element.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
      resetPicture(element.asInlineImage())
    }
  }
} 

function resetAll() {
  let document = DocumentApp.getActiveDocument()
  let body = document.getBody()
  let pictures = body.getImages()
  for (i = 0; i < pictures.length; i++) {
    if (pictures[i].getAltDescription() !== null && pictures[i].getAltDescription().includes("§") !== null) {
      resetPicture(pictures[i])
    }
  }
}

function resetPicture(picture) {
  let textelement = picture.getParent().insertText(picture.getParent().getChildIndex(picture), picture.getAltDescription())
  picture.removeFromParent()
}

function showSidebar() {
  let document = DocumentApp.getActiveDocument()
  let bookmarks = document.getBookmarks()
  console.log(bookmarks)
  for (bi = 0; bi < bookmarks.length; bi++) {
    bookmarks[bi].remove()
  }
  let cursor = document.getCursor()
  bookmark = document.addBookmark(cursor)
  if (bookmark.getPosition().getElement().getType() == DocumentApp.ElementType.PARAGRAPH && bookmark.getPosition().getElement().asParagraph().getNumChildren() === 0) {
    PropertiesService.getDocumentProperties().setProperty("sidebartext", "")
    let html = HtmlService.createHtmlOutputFromFile("sidebar").setTitle("Sidebar").setWidth(300)
    DocumentApp.getUi().showSidebar(html)
  } else {
    bookmark.remove()
    let ui = DocumentApp.getUi()
    ui.alert("Place the cursor on an empty paragraph.")
  }
}

function sidebarUpdated(newtext) {
  PropertiesService.getDocumentProperties().setProperty("sidebartext", newtext)
  let document = DocumentApp.getActiveDocument()
  let bookmarks = document.getBookmarks()
  let bookmark = bookmarks[bookmarks.length - 1]
  let paragraph = bookmark.getPosition().getElement().asParagraph()
  let imagefetch = UrlFetchApp.fetch(encodeURI("https://latex.codecogs.com/png.download?" + newtext), {"muteHttpExceptions": true})
  console.log(encodeURI("https://latex.codecogs.com/png.download?" + newtext))
  if (imagefetch.getResponseCode() !== 200) {
    return
  }
  paragraph.clear()
  let image = imagefetch.getBlob()
  let picture = paragraph.insertInlineImage(0, image)
  picture.setAltDescription("§" + newtext + "§")
  picture.setLinkUrl(encodeURI("www.codecogs.com/eqnedit.php?latex=" + newtext))
  let size = parseFloat(PropertiesService.getDocumentProperties().getProperty("bigsize")) * 1.25
  picture.setWidth(picture.getWidth() * (size / picture.getHeight()))
  picture.setHeight(size)
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
}

function removeBookmark() {
  let document = DocumentApp.getActiveDocument()
  let bookmarks = document.getBookmarks()
  if (bookmarks.length > 0) {
    let bookmark = bookmarks[bookmarks.length - 1]
    let paragraph = bookmark.getPosition().getElement().asParagraph()
    if (paragraph.getNumChildren() > 1) {
      let newtext = PropertiesService.getDocumentProperties().getProperty("sidebartext")
      let imagefetch = UrlFetchApp.fetch(encodeURI("https://latex.codecogs.com/png.download?" + newtext), {"muteHttpExceptions": true})
      console.log(encodeURI("https://latex.codecogs.com/png.download?" + newtext))
      if (imagefetch.getResponseCode() !== 200) {
        bookmark.remove()
        return
      }
      paragraph.clear()
      let image = imagefetch.getBlob()
      let picture = paragraph.insertInlineImage(0, image)
      picture.setAltDescription("§" + newtext + "§")
      picture.setLinkUrl(encodeURI("www.codecogs.com/eqnedit.php?latex=" + newtext))
      let size = parseFloat(PropertiesService.getDocumentProperties().getProperty("bigsize")) * 1.25
      picture.setWidth(picture.getWidth() * (size / picture.getHeight()))
      picture.setHeight(size)
    }
    for (bi = 0; bi < bookmarks.length; bi++) {
      bookmarks[bi].remove()
    }
  }
}

function fixSidebarImage(newtext) {
  let document = DocumentApp.getActiveDocument()
  let bookmarks = document.getBookmarks()
  if (bookmarks.length > 0 && bookmarks[bookmarks.length - 1].getPosition().getElement().asParagraph().getNumChildren() > 1) {
    if (newtext === false) {
      let newtext = PropertiesService.getDocumentProperties().getProperty("sidebartext")
    }
    let bookmark = bookmarks[bookmarks.length - 1]
    let paragraph = bookmark.getPosition().getElement().asParagraph()
    let imagefetch = UrlFetchApp.fetch(encodeURI("https://latex.codecogs.com/png.download?" + newtext), {"muteHttpExceptions": true})
    console.log(encodeURI("https://latex.codecogs.com/png.download?" + newtext))
    if (imagefetch.getResponseCode() !== 200) {
      return
    }
    paragraph.clear()
    let image = imagefetch.getBlob()
    let picture = paragraph.insertInlineImage(0, image)
    picture.setAltDescription("§" + newtext + "§")
    picture.setLinkUrl(encodeURI("www.codecogs.com/eqnedit.php?latex=" + newtext))
    let size = parseFloat(PropertiesService.getDocumentProperties().getProperty("bigsize")) * 1.25
    picture.setWidth(picture.getWidth() * (size / picture.getHeight()))
    picture.setHeight(size)
  }
}
