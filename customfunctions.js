
/**
* Returns a Sentiment score for the passed in text
* @customfunction
* @param connection connection
* @param pov1 period1
* @param pov2 period2
* @helpUrl https://sharepoint.contoso.com/sentiment-faq
*/
async function hsGetValue(connection, pov1, pov2) {

  return storeData[connection];
}
CustomFunctions.associate("HSGETVALUE", hsGetValue);


/** 
 * Returns Sentiment Score
 * @customfunction 
 * @param text string to score
 */
async function happy(text)
{
  return sentiment[text];
}
CustomFunctions.associate("CONTOSO.HAPPYSCORE", happy);


//Store Mock Data
var storeData =
{
    'San Jose Store': 2343,
    'Palo Alto Store': 23432,
    'Ventura Mall Pop Up': 3423,
    'San Francisco Flagship': 3432,
    'Oakland Store': 234,
    'Napa Pop up' : 4324,
    'Cupertino Mall': 2,
    'Marin County Pop Up': 324234,
    'Sacramento Flagship': 342223,
    'Tahoe Casino Popup': 342334,
}


//Sentiment Data
var sentiment =
{
    'This was a great product': 0.95994710922241211,
    'No comment provided.': 0.752673923969268,
    'I like the free wine they offer.  #VinoLife': 0.936662673950195,
    'Wine was delicious': 0.994851231575012,
    'Contoso is the best': 0.994065046310424,
    'YES!': 0.833590745925903,
    'My bike broke and I had to take it back. It was not that fun, but the team was supportive.': 0.0217808485031128,
    'Family feel.': 0.263071537017822,
    'Got a surfboard and it was pretty cool.': 0.994350552558898,
    'hello': 0.921956896781921,
    'This was not that great.': 0.126580029726028,
    'My bike was stolen and Palo Alto Store gave me a new one for free!': 0.227712720632553,
    'So much free stuff': 0.997742295265197,
    'Great': 0.973930835723876,
    'I spent over $100 here. I would definitly shop here again, there is soo much good stuff.': 0.230516254901886,
    'Great products, I was walking buy and decided to hop on in.': 0.900394082069396,
    'Jim Was great': 0.994851231575012,
    'I would recommend the products here to anyone.': 0.918919205665588,
    'Vijay was awesome!': 0.999632835388183,
    'I waited to get help but the guy was too busy….oh well': 0.222730159759521,
    'Was too crazy in there.': 0.726198196411132,
    'Lines were too long. They need more cashiers.': 0.156664729118347,
    'This was a great product.': 0.959947109222412,
    'All the bells and whistles': 0.828923165798187,
    'Pricier than the other locations': 0.0908998250961304,
    'expensive': 0.0460699796676636,
    'Katie was super helpful.': 0.994851231575012,
    'Good': 0.973930835723876,
    'Great, could be better': 0.988386511802673,
}



/*
// The following function is an example that returns a string:
function TRANSLATE(text, langLocale) {
	return new Promise(function(resolve) {
			var xhr = new XMLHttpRequest();
			var textStr = encodeURIComponent(text);
			var localeStr = encodeURIComponent(langLocale);

			var url = "https://excelcf-demo-api.azurewebsites.net/api/translate?code=F69Va5ojUPvfnat9udiM8OpEcScy/oK3bV8/wBYW8OXlypR3nyV/AA==&name=" + 
				textStr + "&locale=" + localeStr;

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


// Use the IEXT API to return the current stock prices for the existing ticker symbol
function STOCKPRICE(ticker)
{
	return new Promise(
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

function STOCKPRICESTREAM(ticker, caller){
	var result = 0;

	//return every second
    setInterval(function(){

		var xhr = new XMLHttpRequest();
		var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
	
		//add handler for xhr
		xhr.onreadystatechange = function() {
			if (xhr.readyState == XMLHttpRequest.DONE) {
		
				//return result back to Excel
				caller.setResult(xhr.responseText);
			}
		}

    	//make request
		xhr.open('GET', url, true);
		xhr.send();    

    }, 1000);
}
*/
