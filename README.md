# MarketplaceScape
A python script to parse *downloaded* facebook marketplace search result listings to track price changes, allowing the user to haggle more effectively knowing the seller's price history. The listings should be the result of a search and the results are saved in a local Excel file.

Because Facebook makes it very apparent that they do not want users scraping their live site, this scripts works with pre downloaded marketplace html webpages that do not contain users' personal data.

## How It Works
- Search for your desired item in the market place search bar.
- Scroll to the end of relevant results
- Right click on a blank part of the page
- Click "save-as"
- Save the file in the same directory as this program

Then run the program. 
It will check for an existing Excel file of listing data and not offer to use existing urls to go and fetch new prices because that is against the rules.
If there is no existing Excel file, it will create a new one.
If you decline or if there is no local Excel file, it will ask if you want to use custom keywords to narrow the results to only include those with the keywords in the title. If you say yes, it will look for a local Excel file of results and read any saved keywords out of the file to ask if you would like to use those. If no, you can enter all new keywords. If yes, just press enter to begin the listing parsing.

The script will then loop through every listing in the html data and check if that listing is already in the Excel file. If it is not in the file, it will add the title, price, town, url, and google map url to the town. If it is, it will compare prices and if it changed, it will update the price, color that cell red, and add the old price with current time stamp to the end of the row.

Next it will loop through the rows of the Excel file and check if any of the html listings are in each row. If there is no match, the Excel listing is out of date and is deleted.

To update your Excel data, simply re-download the webpage and re-run the program providing the appropriate responses in the terminal window.
