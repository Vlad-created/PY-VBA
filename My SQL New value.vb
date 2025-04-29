-- Create a new table to store Sales amount, VAT, and Net value
CREATE TABLE SalesVATNet (
    SalesAmount DECIMAL(10, 2),
    VAT DECIMAL(10, 2),
    NetValue DECIMAL(10, 2)
);

-- Insert Sales amount and VAT from SourceTable into SalesVATNet table
-- Calculate Net value as Sales amount minus VAT
INSERT INTO SalesVATNet (SalesAmount, VAT, NetValue)
SELECT SalesAmount, VAT, SalesAmount - VAT
FROM SourceTable;