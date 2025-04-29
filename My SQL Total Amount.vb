-- Create a general sales table to combine sales from 4 different tables
CREATE TABLE GeneralSales (
    SalesPerson VARCHAR(255),
    TotalSales DECIMAL(10, 2)
);

-- Insert total sales from Table1
INSERT INTO GeneralSales (SalesPerson, TotalSales)
SELECT SalesPerson, SUM(SalesAmount)
FROM Table1
GROUP BY SalesPerson;

-- Insert total sales from Table2
INSERT INTO GeneralSales (SalesPerson, TotalSales)
SELECT SalesPerson, SUM(SalesAmount)
FROM Table2
GROUP BY SalesPerson;

-- Insert total sales from Table3
INSERT INTO GeneralSales (SalesPerson, TotalSales)
SELECT SalesPerson, SUM(SalesAmount)
FROM Table3
GROUP BY SalesPerson;

-- Insert total sales from Table4
INSERT INTO GeneralSales (SalesPerson, TotalSales)
SELECT SalesPerson, SUM(SalesAmount)
FROM Table4
GROUP BY SalesPerson;

-- Select the most productive Sales person who reached amount of 100k
SELECT SalesPerson, TotalSales
FROM GeneralSales
WHERE TotalSales >= 100000
ORDER BY TotalSales DESC;