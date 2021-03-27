# ib_insync_XL
The purpose of this project is to enable others who use Interactive Brokers to write customized trading applications in Python. Interactive Brokers API can be difficult to use, but ib_insync encapsulates the API in an easy-to-use, powerful wrapper. This project goes a step further and ties in Excel, which makes trading significantly easier.

Once the user has come up with a trading strategy, the sample code provided in this project can then be used and extended to write custom algorithms.

Note that:
- All code is written on Macos, due to which some xlWings functionality such as UDFs is not available
- Excel tables are setup in specific formats which much be maintained for the application to work correctly
- For now, the setup works for Stocks and Options only
