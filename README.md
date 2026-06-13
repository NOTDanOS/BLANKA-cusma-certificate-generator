# BLANKA-cusma-certificate-generator
CUSMA Generator for startup company Blanka.

A Google Apps Script tool built during my internship at Blanka (Vancouver, BC) to automate CUSMA export documentation for US-bound shipments.

Background
Our team was manually preparing CUSMA certificates for every US-bound shipment — copying compliance information product by product into a document. For orders with many SKUs sharing the same CUSMA classification, this was both time-consuming and inconsistent, creating real risk for customs delays.
I built this tool to eliminate that process entirely.
How It Works

User pastes SKUs into column B of the CUSMA sheet
Script cross-references a master product inventory database to retrieve each SKU's compliance data — goods description, HTS tariff code, origin criterion, country of origin, and producer name
Flags any SKUs that are missing, unrecognized, or have incomplete data before writing anything
Populates the CUSMA certificate automatically — ready to download and send

Versions
V1 (cusma_generator_v1.gs)

Maximum 8 SKUs per document
Auto-populated a separate producer section with full manufacturer details (address, phone, email) pulled from a second database
Included a custom spreadsheet menu via onOpen()

V2 (cusma_generator_v2.gs) (current)

Expanded SKU capacity to 20 per document — resolving the bottleneck on large orders
Simplified producer handling — name populated inline per product row rather than in a separate section
Cleaner template structure consolidated from row 19

Impact

Restored daily US shipment capacity from 10–20 to 30–40 orders
Reduced per-order shipping costs by ~50% compared to non-compliant shipments

Notes

Built with Google Apps Script (JavaScript)
Retrieves data from a private company spreadsheet — the sheet ID is included for documentation purposes only and the script will not run without access to the original data source
