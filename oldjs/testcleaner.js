let text = `{ "SALES CONTACT":"James", "WECHAT":"0086-13626696200", "PRODUCT REAL DESCRIPTION":"About material to "sewing" version ", "WECHAT":"0086-13626696200"}`;

function cleanText(dirtyText) {
  // Step 1: Extract substrings between a colon and a comma or a closing curly bracket
  const regex = /:\s*([^,}]*)[,\}]/g;
  let match;
  let cleanedText = dirtyText;

  while ((match = regex.exec(dirtyText)) !== null) {
    if ((match[1].match(/"/g) || []).length > 2) {
      console.log("mal! ", match[1]);
      const split = match[1].split('"');
      let singleQuote = `${split.join("'")}`;
      singleQuote = `"${singleQuote.slice(1, -1)}"`;
      console.log("singleQuote", singleQuote);
      cleanedText = cleanedText.replace(match[1], singleQuote);
    }
  }
  console.log("cleanedText", cleanedText);

  return cleanedText;
}
cleanText(text);
/* module.exports = {
  cleanText,
}; */
