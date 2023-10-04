# VBA-challenge
In this challenge we used VBA scripting to analyse generated stock market data.

I started using cycle "For each" which allows us to loop the data for the years that have the excel file. This means that if next year we add 2024, the script will keep working succesfully.

I used constants for the columns and rows positions so thats its user maintenance friendly

A script that loops (do while) through all stocks and outputs the total yearly change, percent change, and stock volume for each ticker at the end of each month was created.

To achieve this, I took the open value in the begining of each ticker year and the close value at the end of each ticker year to calculate the yearly percent change as follow:
(((lClose - lOpen) * 100) / lOpen) / 100

The total stock volume was accumalating meanwhile it was going through each ticker. When the ticker changed, it dumped the data in the TotalsCells, just as the yearly percent change calculation mentioned in the previous paragraph.

To save the greatest values I defined variables for the amount and names with the prefix "greatest". Then compared one variable with the previuos variable each time we print the totals keeping the greatest value, and finally printed the greatest values in the end of the loop before changing the sheet.

According to our analysis,
for 2018, ticker THB had the greatest percent increase, RKS had the greatest decrease, and QKN the most volume.
for 2019, ticker RYU had the greatest percent increase, RKS had the greatest decrease, and ZQD the most volume.
for 2020, ticker YDI had the greatest percent increase, VNG had the greatest decrease, and QKN the most volume.

RKS had the greatest percent decrease for two consequtive years, 2018 and 2019.
Buying multiple PUT Options for RKS in 2017 would have been a very profitable investment.

       
