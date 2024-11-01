This program is designed to monitor and adjust product prices to stay competitive. 
It checks prices listed by other sellers for the same products and, if a competitor's price is higher than a set minimum, the program reduces its own price to remain lower. 
This approach allows the business to stay more attractive to buyers by always offering a slightly lower price than competitors without dropping below a specified threshold.

![Screenshot](https://github.com/parserKASPI/blob/main/Screenshot_232.png)

Here's how the program works:

Selenium setup – The program configures Selenium, a tool for web automation, to interact with the website in a headless mode (without a visible browser window). This setup includes options to reduce logging noise and ensure smooth operation without interruptions.

Logging in – Using login details stored in an Excel file, the program logs in to the website’s account portal. This lets it access features like product management, where it can update prices.

Getting competitor’s product price – The program navigates to the competitor’s product page to find and retrieve the current price. It then compares this price with the set minimum price to decide whether an update is needed.

Updating own price – If the competitor’s price is higher than the minimum, the program updates its own price to stay below it by a certain amount (defined by the user). This keeps the business’s price slightly more attractive.

Main loop – The program goes through all products listed in the spreadsheet, repeating the process for each one. After checking all items, it waits for a set interval before the next cycle, ensuring prices remain competitive over time.

