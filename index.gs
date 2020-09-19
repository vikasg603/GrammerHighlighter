function countWords(str) {
  var matches = str.match(/[\w\d\’\'-]+/gi);
  return matches ? matches.length : 0;
}

function App() {
 
  var body = DocumentApp.getActiveDocument().getBody();
  
  const RegexForVerbs = '(?i)([ \t\r\n\f]|^)(Is|Are|Able|Am|Be|Being|Become|Becomes|Became|Becoming|Been|Begin|Beginning|Has|Have|Had|Was|Were)([ \t\r\n\f]|$)';
  const RegexForPrepositions = '(?i)([ \t\r\n\f]|^)(About|Above|Across|After|Around|At|As|Before|Behind|Below|Between|Beside|By|Of|From|For|Into|In|Like|Off|On|Onto|Out|With|Within|Without|Under|Up|Upon|Until|That|Than|To|Toward|Through|Towards)([ \t\r\n\f]|$)';
  
  const RegexForWhiteSpaces = /[ \t\r\n\f]/;
  const ColorForVerbs = '#FCFC00';
  const ColorForPrepositions = '#00ff00';
  
  var VerbsCount = 0;
  var PrepositionsCount = 0;
  
  const HighlightedWords = {};
  
  //First Deleting Old report 
  var Text = body.getText();
  var Start = Text.indexOf('======BEGIN REPORT======');
  if(Start >= 0) {
    var End = Text.length - 1;
    body.editAsText().deleteText(Start - 1, End); // -1 as to add new line too
  }
  
  //As We have updated our text body, so retriving new one
  var body = DocumentApp.getActiveDocument().getBody();
  //Finding Verbs
  var foundElement = body.findText(RegexForVerbs);
  
  while (foundElement != null) {
    // Get the text object from the element
    var foundText = foundElement.getElement().asText();
    
    // Where in the element is the found text?
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    
    //Plus 1 as we need to include last element
    var foundTextAsString = foundText.getText().slice(start, end + 1);
    
    //Removing any spaces if it's there at the end of word
    if(RegexForWhiteSpaces.test(foundTextAsString.charAt(foundTextAsString.length - 1))) {
      end = end - 1;
    }
    
    //Removing any spaces if it's there at the beginning of word
    if(RegexForWhiteSpaces.test(foundTextAsString.charAt(0))) {
      start = start + 1;
    }
    
    //Now getting Word in lower case
    foundTextAsString = foundText.getText().slice(start, end + 1).toLowerCase();
    
    
    // Setting highlighted word as Bold
    foundText.setBold(start, end, true);
    
    // Change the background color to ColorForVerbs
    foundText.setBackgroundColor(start, end, ColorForVerbs);
    
    //Incrementing count
    VerbsCount++;
    
    
    //Incrementing HighlightWords in HighlightWords object
    isFinite(HighlightedWords[foundTextAsString]) ? (HighlightedWords[foundTextAsString] = HighlightedWords[foundTextAsString] + 1) : (HighlightedWords[foundTextAsString] = 1);
    
    // Find the next match
    foundElement = body.findText(RegexForVerbs, foundElement);
  }
  //Verbs ends here
  
  //Finding Prepositions
  foundElement = body.findText(RegexForPrepositions);
  
  while (foundElement != null) {
    // Get the text object from the element
    var foundText = foundElement.getElement().asText();
    
    // Where in the element is the found text?
    var start = foundElement.getStartOffset(); 
    var end = foundElement.getEndOffsetInclusive();
    
    //Plus 1 as we need to include last element
    var foundTextAsString = foundText.getText().slice(start, end + 1);
    
    
    //Removing any spaces if it's there at the end of word
    if(RegexForWhiteSpaces.test(foundTextAsString.charAt(foundTextAsString.length - 1))) {
      end = end - 1;
    }
    
    //Removing any spaces if it's there at the beginning of word
    if(RegexForWhiteSpaces.test(foundTextAsString.charAt(0))) {
      start = start + 1;
    }
    
    //Now getting Word in lower case
    foundTextAsString = foundText.getText().slice(start, end + 1).toLowerCase();
    
    // Set Bold
    foundText.setBold(start, end, true);
    
    // Change the background color to ColorForPrepositions
    foundText.setBackgroundColor(start, end, ColorForPrepositions);
    
    //Incrementing count
    PrepositionsCount++;
    
    //Incrementing HighlightWords in HighlightWords object
    isFinite(HighlightedWords[foundTextAsString]) ? (HighlightedWords[foundTextAsString] = HighlightedWords[foundTextAsString] + 1) : (HighlightedWords[foundTextAsString] = 1);
    
    
    // Find the next match
    foundElement = body.findText(RegexForPrepositions, foundElement);
  }
  //Prepositions ends here
  
  
  //Inserting Report at the end.
  var Text = body.getText();
  var TotalWords = countWords(Text);
 
  const VerbCountPercentage = ((VerbsCount/TotalWords)*100).toFixed(2);
  const PrepositionsCountPercentage = ((PrepositionsCount/TotalWords)*100).toFixed(2);
  
  body.appendParagraph('======BEGIN REPORT======');
  
  //Hightlighted Words From HighlightedWords object;
  Object.keys(HighlightedWords).forEach(function (item) {
    body.appendParagraph("Highlighted '" + item + "' " + HighlightedWords[item] + " times.");
  });
  
  
  body.appendParagraph('');
  
  body.appendParagraph('Yellow highlights indicate weak verbs.');
  body.appendParagraph('Green highlights indicate prepositions.');
  body.appendParagraph('Red highlights indicate expletives.');
  body.appendParagraph('Edit colorful sentences.');
  body.appendParagraph('Weak verb fraction = ' + VerbCountPercentage + '%. Strive for < 0.5%.');
  body.appendParagraph('Preposition fraction = ' + PrepositionsCountPercentage + '%. Strive for < 10%.');
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Run', 'App')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}
