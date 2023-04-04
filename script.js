/**
 * Convert a HTML code into it's correspondent BBCode
 * @param {html} input the Cell to convert.
 * @return The converted BBCode text.
 * @customfunction
*/
function HTMLTOBBCODE(html) {
  // Replace HTML tags with BBCode equivalents
  html = html.replace(/<b(.*?)>(.*?)<\/b>/gmi, '[b]$2[/b]');
  html = html.replace(/<strong(.*?)>(.*?)<\/strong>/gmi, '[b]$2[/b]');
  html = html.replace(/<i(.*?)>(.*?)<\/i>/gmi, '[i]$2[/i]');
  html = html.replace(/<em(.*?)>(.*?)<\/em>/gmi, '[i]$2[/i]');
  html = html.replace(/<u(.*?)>(.*?)<\/u>/gmi, '[u]$2[/u]');
  html = html.replace(/<s(.*?)>(.*?)<\/s>/gmi, '[s]$2[/s]');
  html = html.replace(/<blockquote(.*?)>(.*?)<\/blockquote>/gmi, '[quote]$2[/quote]');
  html = html.replace(/<img(.*?)src="(.*?)"(.*?)>/gmi, '[img]$2[/img]');
  html = html.replace(/<a(.*?)href="(.*?)"(.*?)>(.*?)<\/a>/gmi, '[url=$2]$4[/url]');
  html = html.replace(/<ul(.*?)>/gmi, '[list]');
  html = html.replace(/<\/ul>/gmi, '[/list]');
  html = html.replace(/<li(.*?)>(.*?)<\/li>/gmi, '[*]$2');
  html = html.replace(/<ol(.*?)>/gmi, '[list=1]');
  html = html.replace(/<\/ol>/gmi, '[/list]');
  html = html.replace(/<table(.*?)>(.*?)<\/table>/gmi, '[table]$2[/table]');
  html = html.replace(/<tr(.*?)>(.*?)<\/tr>/gmi, '[tr]$2[/tr]');
  html = html.replace(/<td(.*?)>(.*?)<\/td>/gmi, '[td]$2[/td]');
  html = html.replace(/<th(.*?)>(.*?)<\/th>/gmi, '[th]$2[/th]');
  html = html.replace(/<h1(.*?)>(.*?)<\/h1>/gmi, '[size=6]$2[/size]');
  html = html.replace(/<h2(.*?)>(.*?)<\/h2>/gmi, '[size=5]$2[/size]');
  html = html.replace(/<h3(.*?)>(.*?)<\/h3>/gmi, '[size=4]$2[/size]');
  html = html.replace(/<h4(.*?)>(.*?)<\/h4>/gmi, '[size=3]$2[/size]');
  html = html.replace(/<h5(.*?)>(.*?)<\/h5>/gmi, '[size=3]$1[/size]');
  html = html.replace(/<h6(.*?)>(.*?)<\/h6>/gmi, '[size=3]$1[/size]');

  // Replace <p></p>, <span>, <br> and other tags with nothing
  html = html.replace(/<p(.*?)>(.*?)<\/p>/gmi, '$2\n');
  html = html.replace(/<o:p(.*?)>(.*?)<\/o:p>/gmi, '$2\n'); //office namespace
  html = html.replace(/<p>/gi, '');
  html = html.replace(/<\/p>/gi, '\n\n');
  html = html.replace(/<br \/>/gi, '\n');
  html = html.replace(/<br(.*?)>/gi, '\n');
  html = html.replace(/<span(.*?)>(.*?)<\/span>/gmi, '$2');
  html = html.replace(/<span(.*?)>/gi, '');
  html = html.replace(/<\/span>/gi, '');
  html = html.replace(/<div(.*?)>(.*?)<\/div>/gmi, '$2');
  html = html.replace(/<div(.*?)>/gi, '');
  html = html.replace(/<\/div>/gi, '');
  html = html.replace(/<font(.*?)>/gi, '');
  html = html.replace(/<\/font>/gi, '');
  html = html.replace(/<tbody(.*?)>/gi, '');
  html = html.replace(/<\/tbody>/gi, '');

  // Replace dashes and dots characters
  html = html.replace(/&mdash;/gi, '—'); 
  html = html.replace(/&ndash;/gi, '–'); 
  html = html.replace(/&middot;/gi, '·'); 
  html = html.replace(/&bull;/gi, '•'); 

  //Unusual characters
  html = html.replace(/&#10003;/gi, '✓');

  //Converts hexadecimal HTML entities
  html = html.replace(/&#(\d+);/g, ((match, dec) => `${String.fromCharCode(dec)}`));

  // Replace HTML entities with their text equivalents
  html = html.replace(/&nbsp;/gi, ' ');
  html = html.replace(/&lt;/gi, '<');
  html = html.replace(/&gt;/gi, '>');
  html = html.replace(/&quot;/gi, '"');
  html = html.replace(/&apos;/gi, "'");
  html = html.replace(/&lsquo;/gi, "‘");
  html = html.replace(/&rsquo;/gi, "’");
  html = html.replace(/&sbquo;/gi, "‚");
  html = html.replace(/&ldquo;/gi, "“");
  html = html.replace(/&rdquo;/gi, "”");
  html = html.replace(/&ccedil;/g, 'ç');
  html = html.replace(/&Ccedil;/g, 'Ç');
  html = html.replace(/&aacute;/g, 'á');
  html = html.replace(/&agrave;/g, 'à');
  html = html.replace(/&Aacute;/g, 'Á');
  html = html.replace(/&eacute;/g, 'é');
  html = html.replace(/&Eacute;/g, 'É');
  html = html.replace(/&iacute;/g, 'í');
  html = html.replace(/&Iacute;/g, 'Í');
  html = html.replace(/&oacute;/g, 'ó');
  html = html.replace(/&Oacute;/g, 'Ó');
  html = html.replace(/&uacute;/g, 'ú');
  html = html.replace(/&Uacute;/g, 'Ú');
  html = html.replace(/&atilde;/g, 'ã');
  html = html.replace(/&Atilde;/g, 'Ã');
  html = html.replace(/&otilde;/g, 'õ');
  html = html.replace(/&Otilde;/g, 'Õ');
  html = html.replace(/&ntilde;/g, 'ñ');
  html = html.replace(/&Ntilde;/g, 'Ñ');
  html = html.replace(/&acirc;/g, 'â');
  html = html.replace(/&Acirc;/g, 'Â');
  html = html.replace(/&ecirc;/g, 'ê');
  html = html.replace(/&Ecirc;/g, 'Ê');
  html = html.replace(/&icirc;/g, 'î');
  html = html.replace(/&Icirc;/g, 'Î');
  html = html.replace(/&ocirc;/g, 'ô');
  html = html.replace(/&Ocirc;/g, 'Ô');
  html = html.replace(/&deg;/gi, '°');
  html = html.replace(/&ordf;/gi, 'ª');
  html = html.replace(/&ordm;/gi, 'º');
  html = html.replace(/&oslash;/gi, 'ø');
  html = html.replace(/&Oslash;/gi, 'Ø');
  html = html.replace(/&amp;/gi, '&');
  html = html.replace(/&nbsp;/gi, ' ');

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
