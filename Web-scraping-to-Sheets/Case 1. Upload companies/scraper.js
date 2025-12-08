function scraper() {
  
  // Get sheet named 'Atlanta'
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Atlanta');
  // Read base URLs from row 1
  const urlDomain1 = ss.getRange(1, 1).getValue();
  const urlDomain = ss.getRange(1, 2).getValue();

  // Process the first URL (usually the catalog’s main page)
  let url1 = urlDomain1;
  let count1 = 0;
  // Send HTTP request to url1
  let response1 = UrlFetchApp.fetch(url1);
  // Write response status code to A2
  ss.getRange(2, 1).setValue(response1.getResponseCode());
  let html = response1.getContentText();
  let p1 = 0;

    // Loop: find all blocks with class "provider_card" and process them
    while (true) {    

      // Extract next <div>...</div> provider card block
      let out = getBlock(html, 'div', html.indexOf('class="provider_card"', p1));
      let block = out[0];
      p1 = out[1] + 1;
      if (p1 == 0) break; // stop if no more blocks found

      // Extract provider name (<a class="provider_name">)
      let title = getBlock(block, 'a', block.indexOf('class="provider_name"'))[0];
      // Write title into column 2
      ss.getRange(4 + 1 * count1, 2).setValue(title).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        // Extract location (div provider_info__item location)
        let location = getBlock(block, 'div', block.indexOf('class="provider_info__item location"'))[0];
        location = location.split('</span>').map(item => item.split('<span>')[1]);
        ss.getRange(4 + 1 * count1, 3).setValue(location).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        // Extract company size
        let size = getBlock(block, 'div', block.indexOf('class="provider_info__item size"'))[0];
        size = size.split('</span>').map(item => item.split('<span>')[1]);
        ss.getRange(4 + 1 * count1, 4).setValue(size).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        // Extract list of services
        let services = getBlock(block, 'div', block.indexOf('class="provider_details__wrap service"'))[0];
        services = services.split('</span>').map(item => item.split('>')[1]);
        services.pop();
        services = services.join(', ');
        ss.getRange(4 + 1 * count1, 5).setValue(services).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        // Extract link from provider header block
        let link = getBlock(block, 'div', block.indexOf('class="provider_header__top"'))[0];
        let link2 = getOpenTag(link, 'a', 0);
        link2 = getAttrName(link2, 'href', 0)
        ss.getRange(4 + 1 * count1, 9).setValue(link2).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy

      count1++;

    };

  // Process paginated pages
  let url = urlDomain;
  let count = 0;
  // Pages 1..12
  for (let i = 1; i < 13; i++) {

    if (ss.getRange(53 , 2) != "") {
  
      url = urlDomain + i;
      let response = UrlFetchApp.fetch(url);
      ss.getRange(2, 2).setValue(response.getResponseCode());
      let html = response.getContentText();
      let p = 0;

      while (true) {    

        let out = getBlock(html, 'div', html.indexOf('class="provider_card"', p));
        let block = out[0];
        p = out[1] + 1;
        if (p == 0) break;

        let title = getBlock(block, 'a', block.indexOf('class="provider_name"'))[0];
        ss.getRange(54 + 1 * count, 2).setValue(title).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        let location = getBlock(block, 'div', block.indexOf('class="provider_info__item location"'))[0];
        location = location.split('</span>').map(item => item.split('<span>')[1]);
        ss.getRange(54 + 1 * count, 3).setValue(location).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        let size = getBlock(block, 'div', block.indexOf('class="provider_info__item size"'))[0];
        size = size.split('</span>').map(item => item.split('<span>')[1]);
        ss.getRange(54 + 1 * count, 4).setValue(size).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        let services = getBlock(block, 'div', block.indexOf('class="provider_details__wrap service"'))[0];
        services = services.split('</span>').map(item => item.split('>')[1]);
        services.pop();
        services = services.join(', ');
        ss.getRange(54 + 1 * count, 5).setValue(services).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        let link = getBlock(block, 'div', block.indexOf('class="provider_header__top"'))[0];
        let link2 = getOpenTag(link, 'a', 0);
        link2 = getAttrName(link2, 'href', 0)
        ss.getRange(54 + 1 * count, 9).setValue(link2).setHorizontalAlignment('left').setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        count++;
      };
    }

  };
}

// Add menu button when sheet opens
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Script Menu')
  .addItem('Run script','scraper')
  .addToUi();
}

// Utility: extract attribute value (e.g., href)
function getAttrName(html, attr, i) {
  let idxStart = html.indexOf(attr +'=' , i);
  if (idxStart == -1) return "Can't find attr " + attr + ' !';
  idxStart = html.indexOf('"' , idxStart) + 1;
  let idxEnd = html.indexOf('"' , idxStart);
  return html.slice(idxStart,idxEnd).trim();
}

// Utility: returns opening tag like <a ...>
function getOpenTag(html, tag, idxStart) {
  let openTag = '<' + tag;
  let lenOpenTag = openTag.length;
  if (html.slice(idxStart, idxStart + lenOpenTag) != openTag) {
    idxStart = html.lastIndexOf(openTag, idxStart);
    if (idxStart == -1) return "Can't find openTag " + openTag + ' !';
  };
  let idxEnd = html.indexOf('>', idxStart) + 1;
  if (idxStart == -1) return "Can't find closing bracket '>' for openTag!";
  return html.slice(idxStart,idxEnd).trim();
}

// Utility: delete tag block including its content
function deleteBlock(html, tag, idxStart) { // delete opening & closing tag and info between them
  let openTag = '<' + tag;
  let lenOpenTag = openTag.length;
  let closeTag = '</' + tag + '>';
  let lenCloseTag = closeTag.length;
  let countCloseTags = 0;
  let iMax = html.length;
  let idxEnd = 0;
  // where we are?
  if (html.slice(idxStart, idxStart + lenOpenTag) != openTag) {
    idxStart = html.lastIndexOf(openTag, idxStart);
    if (idxStart == -1) return ["Can't to find openTag " + openTag + ' !', -1];
  };
  // begin loop after openTag
  let i = html.indexOf('>') + 1;
  
  while (i <= iMax) {
    i++;
    if (i === iMax) {
      return ['Could not find closing tag for ' + tag, -1];
    };
    let carrentValue = html[i];
    if (html[i] === '<'){
      let closingTag = html.slice(i, i + lenCloseTag);
      let openingTag = html.slice(i, i + lenOpenTag);
      if (html.slice(i, i + lenCloseTag) === closeTag) {
        if (countCloseTags === 0) {
          idxEnd = i + lenCloseTag;
          break;
        } else {
          countCloseTags -= 1;
        };
      } else if (html.slice(i, i + lenOpenTag) === openTag) {
        countCloseTags += 1;
      };
    };
  };
  return (html.slice(0, idxStart) + html.slice(idxEnd, iMax)).trim();
}

function getBlock(html, tag, idxStart) {  // <tag .... > Block </tag>
  let openTag = '<' + tag;
  let lenOpenTag = openTag.length;
  let closeTag = '</' + tag + '>';
  let lenCloseTag = closeTag.length;
  let countCloseTags = 0;
  let iMax = html.length;
  let idxEnd = 0;
  // where we are?
  if (html.slice(idxStart, idxStart + lenOpenTag) != openTag) {
    idxStart = html.lastIndexOf(openTag, idxStart);
    if (idxStart == -1) return ["Can't to find openTag " + openTag + ' !', -1];
  };
  // change start - will start after openTag!
  idxStart = html.indexOf('>', idxStart) + 1;
  let i = idxStart;
  
  while (i <= iMax) {
    i++;
    if (i === iMax) {
      return ['Could not find closing tag for ' + tag, -1];
    };
    let carrentValue = html[i];
    if (html[i] === '<'){
      let closingTag = html.slice(i, i + lenCloseTag);
      let openingTag = html.slice(i, i + lenOpenTag);
      if (html.slice(i, i + lenCloseTag) === closeTag) {
        if (countCloseTags === 0) {
          idxEnd = i - 1;
          break;
        } else {
          countCloseTags -= 1;
        };
      } else if (html.slice(i, i + lenOpenTag) === openTag) {
        countCloseTags += 1;
      };
    };
  };
  return [html.slice(idxStart,idxEnd + 1).trim(), idxEnd];
}


