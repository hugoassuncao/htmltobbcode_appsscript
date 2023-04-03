/**
 * Convert a HTML code into it's correspondent BBCode
 * @param {html} input the Cell to convert.
 * @return The converted BBCode text.
 * @customfunction
*/
function HTMLTOBBCODE(html) {
  // Replace HTML tags with BBCode equivalents
  html = html.replace(/<b>(.*?)<\/b>/gi, '[b]$1[/b]');
  html = html.replace(/<strong>(.*?)<\/strong>/gi, '[b]$1[/b]');
  html = html.replace(/<i>(.*?)<\/i>/gi, '[i]$1[/i]');
  html = html.replace(/<u>(.*?)<\/u>/gi, '[u]$1[/u]');
  html = html.replace(/<s>(.*?)<\/s>/gi, '[s]$1[/s]');
  html = html.replace(/<img(.*?)src="(.*?)"(.*?)>/gi, '[img]$2[/img]');
  html = html.replace(/<a(.*?)href="(.*?)"(.*?)>(.*?)<\/a>/gi, '[url=$2]$4[/url]');
  html = html.replace(/<ul>/gi, '[list]');
  html = html.replace(/<\/ul>/gi, '[/list]');
  html = html.replace(/<li>(.*?)<\/li>/gi, '[*]$1');
  html = html.replace(/<ol>/gi, '[list=1]');
  html = html.replace(/<\/ol>/gi, '[/list]');
  html = html.replace(/<table(.*?)>(.*?)<\/table>/gi, '[table]$2[/table]');
  html = html.replace(/<tr(.*?)>(.*?)<\/tr>/gi, '[tr]$2[/tr]');
  html = html.replace(/<td(.*?)>(.*?)<\/td>/gi, '[td]$2[/td]');
  html = html.replace(/<th(.*?)>(.*?)<\/th>/gi, '[th]$2[/th]');

  // Replace <p></p> tags with nothing
  html = html.replace(/<p>(.*?)<\/p>/gi, '$1\n\n');

  // Replace dashes characters
  html = html.replace(/&mdash;/gi, '—'); // m-dash
  html = html.replace(/&ndash;/gi, '–'); // n-dash

  // Replace HTML entities with their text equivalents
  html = html.replace(/&nbsp;/gi, ' ');
  html = html.replace(/&lt;/gi, '<');
  html = html.replace(/&gt;/gi, '>');
  html = html.replace(/&quot;/gi, '"');
  html = html.replace(/&apos;/gi, "'");
  html = html.replace(/&ccedil;/gi, 'ç');
  html = html.replace(/&Ccedil;/gi, 'Ç');
  html = html.replace(/&aacute;/gi, 'á');
  html = html.replace(/&Aacute;/gi, 'Á');
  html = html.replace(/&eacute;/gi, 'é');
  html = html.replace(/&Eacute;/gi, 'É');
  html = html.replace(/&iacute;/gi, 'í');
  html = html.replace(/&Iacute;/gi, 'Í');
  html = html.replace(/&oacute;/gi, 'ó');
  html = html.replace(/&Oacute;/gi, 'Ó');
  html = html.replace(/&uacute;/gi, 'ú');
  html = html.replace(/&Uacute;/gi, 'Ú');
  html = html.replace(/&atilde;/gi, 'ã');
  html = html.replace(/&Atilde;/gi, 'Ã');
  html = html.replace(/&otilde;/gi, 'õ');
  html = html.replace(/&Otilde;/gi, 'Õ');

  return html;
}

/**
 * Decode HTML entities in a range - eg.: "SheetName!A1:A", between quotes.
 * @param {range} input the range to decode, eg.: "SheetName!A1:A", between quotes.
 * @return The converted text.
 * @customfunction
*/
function HTMLTOBBCODERANGE(range) {
  let rangeValues = SpreadsheetApp.getActive().getRange(range).getValues();
  var valuesArray = [];
  rangeValues.forEach(function(item,i){
    var itemR = HTMLTOBBCODE(item[0]);
    valuesArray.push(itemR);
  });
  return valuesArray;
}
