//let text = `{ "SALES CONTACT":"James", "WECHAT":"0086-13626696200", "PRODUCT REAL DESCRIPTION":"About this bag we can not do because it"s heat sealed version. if change the outer material to sewing version we can make it. like below the picture.", "SALES CONTACT":"James", "WECHAT":"0086-13626696200", "PRODUCT REAL DESCRIPTIONxxx":"About this bag we can not do because it"s heat sealed version. if change the outer material to sewing version we can make it. like below the picture.", "WECHAT":"0086-13626696200"}`;

function cleanText(dirtyText) {
  // Step 1: Extract substrings between a colon and a comma or a closing curly bracket
  const regex = /:\s*([^,}]*)[,\}]/g;
  let match;
  while ((match = regex.exec(dirtyText)) !== null) {
    if ((match[1].match(/"/g) || []).length > 2) {
      console.log("mal! ", match[1]);
      const split = match[1].split('"');
      let singleQuote = `${split.join("'")}`;
      singleQuote = `"${singleQuote.slice(1, -1)}"`;
      console.log(singleQuote);
      text = dirtyText.replace(match[1], singleQuote);
    }
  }
  console.log(text);
  return text;
}
