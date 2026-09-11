# Inventory Management System
Created an Inventory System,
View the .XLSX file without the '~$'.
"~$DataAnalyst.xlsx" file is a dummy file, not for viewing; view the other one.

This inventory system relies on a few key Excel formulas to keep data dynamic. Here are the formulas used

1. Product Details Lookup (Inventory Sheet):
Used to pull product information from the master product catalog based on the Product ID.

Formula: =XLOOKUP([@[Product ID]], Products[Product ID], Products[Product Name], "Not Found") (Note: The return_array is changed to reference other columns like Supplier or Cost per Unit as needed.)
2. Quantity on Hand (Inventory Sheet):
Calculates current stock by summing all transactions (receipts and sales) for a specific product and site.

Formula: =SUMIFS(Transactions[Quantity], Transactions[Product ID], [@[Product ID]], Transactions[Site], [Site])
3. Stock Value (Inventory Sheet):
Multiplies the unit cost by the quantity currently on hand.

Formula: =[@[Cost per Unit]] * [@[Quantity on Hand]]
4. Reorder Logic (Inventory Sheet):
Determines if an item needs to be restocked based on the defined reorder level.

Formula: =IF([@[Quantity on Hand]] <= [@[Reorder Level]], "Yes", "No")
5. Total Stock Value (Inventory Sheet Header):
Sums the entire stock value column for a high-level report.

Formula: =SUM(Inventory[Stock Value])
