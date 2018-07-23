
// The following function is an example that returns a string:
function TRANSLATE(text, langLocale) {
	return new OfficeExtension.Promise(function(resolve) {
			var xhr = new XMLHttpRequest();
			var url = 'https://dev.office.com'
			
			xhr.onreadystatechange = function() {
				if (xhr.readyState == XMLHttpRequest.DONE) {
					resolve(xhr.responseText);
				}
			}
			xhr.open('GET', url, true);
			xhr.send();
	});
}


function REGEXMATCH(text, regex)
{
	//call the building in string match function which supports regular expressions
	var res =  text.match(regex);

	return (res != null); //TRUE if there's a match, otherwise FALSE
}


function REGEXREPLACE(text, regex, replace_text)
{
	var regexToken = new RegExp(regex);
	var res =  text.replace(regexToken, replace_text);
	return (res != null) ? res : "Error: no results found to replace";
}


// Use the IEXT server to return the current stock prices for the existing ticker symbol
function STOCKPRICE(ticker)
{
   
  return new OfficeExtension.Promise(
	  function(resolve) {
	
		var xhr = new XMLHttpRequest();
		var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
	
		//add handler for xhr
		xhr.onreadystatechange = function() {
			if (xhr.readyState == XMLHttpRequest.DONE) {
				
				//return result back to Excel
				resolve(xhr.responseText);
			}
		}

		//make request
		xhr.open('GET', url, true);
		xhr.send();
});

}