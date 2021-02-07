import ib_insync as ibi
import xlwings as xw

def accountValue(accountNumber, tag: str) -> str:
    """Return account value for the given tag."""
    return next((
        v.value for v in ib.accountValues(accountNumber)
        if v.tag == tag and v.currency == 'BASE'), '')


def getAccountNumberDict():
    return {"Enter your account number here": "paper"}


def getIbConnectionPort(xlSheetName):
    # return the socket port number to use for connecting to TWS
    if xlSheetName == "Paper":
        return 7497
    else:
        return 7496


def closePositions():
    wb = xw.Book.caller()
    callingWorksheet = wb.sheets.active.name

    # Get the Excel cell address where "Positions to Close" table begins
    cellAddress = wb.sheets(callingWorksheet).tables[callingWorksheet + 'PositionsToCloseTable'].data_body_range(1, 1).address

    # Pick up the "Positions to Close" table
    positionsToClose = wb.sheets(callingWorksheet).tables[callingWorksheet + 'PositionsToCloseTable'].data_body_range.options(ndim=2).value
    
    # If the first cell of the "Positions to Close" is empty, raise error
    if positionsToClose[0][0] == None or positionsToClose[0][0].strip() == "":
        raise ValueError ("No Positions to Close")

    # Start a new IB connection
    ib.connect('127.0.0.1', getIbConnectionPort(callingWorksheet), clientId=2)

    # Match each position to be closed with currently held positions, per IB's records
    # # If a match is found, place closing order
    for acccountPosition in ib.positions():
        for position in positionsToClose:
            if acccountPosition.contract.localSymbol.lower().replace(" ", "") == position[0].lower().replace(" ", ""):
                # if current position for a contract is LONG, then set closing action to "SELL"
                # otherwise set closing action to "BUY"
                closingAction = "SELL" if acccountPosition.position > 0 else "BUY"

                # Prepare closing order
                if position[1].upper() == "LMT":
                    # If limit price is not set, then raise error
                    if ibi.util.isNan(position[2]) or position[2] < 0:
                        ib.disconnect()
                        raise ValueError ("Incorrect Limit Price for " + position[0])
                    else:
                        closingOrder = ibi.LimitOrder(closingAction, abs(acccountPosition.position), position[2])
                elif position[1].upper() == "MKT":
                    closingOrder = ibi.MarketOrder(closingAction, abs(acccountPosition.position))
                else:
                    raise ValueError ("Incorrect Order Type for " + position[0])

                trade = ib.placeOrder(acccountPosition.contract, closingOrder)
                assert trade in ib.trades()

    # Disconnect from IB after placing the orders
    ib.disconnect()


def placeOrders():
    wb = xw.Book.caller()
    callingWorksheet = wb.sheets.active.name

    # Get the Excel cell address where "Positions to Close" table begins
    cellAddress = wb.sheets(callingWorksheet).tables[callingWorksheet + 'OrderListTable'].data_body_range(1, 1).address

    # Pick up the "Positions to Close"
    ordersFromXL = wb.sheets(callingWorksheet).tables[callingWorksheet + 'OrderListTable'].data_body_range.options(ndim=2).value
    
    # If the first cell of the "Positions to Close" is empty, raise error
    if ordersFromXL[0][0] == None or ordersFromXL[0][0].strip() == "":
        raise ValueError ("No Orders to Submit")

    # Start a new IB connection
    ib = ibi.IB()
    ib.connect('127.0.0.1', getIbConnectionPort(callingWorksheet), clientId=4)

    # Place orders one at a time
    for order in ordersFromXL:
        # Create the entryOrder object
        if order[2].upper() == "LMT":
            entryOrder = ibi.LimitOrder(order[1], abs(int(order[6])), order[5])
        elif order[2].upper() == "MKT":
            entryOrder = ibi.MarketOrder(order[1], abs(int(order[6])))
        else:
            raise ValueError ("Incorrect Order Type in " + order[0])

        # Create the contract object
        if order[7].upper() == "STK":
            contract = ibi.Stock(order[8], 'SMART', 'USD', primaryExchange='NYSE')
        elif order[7].upper() == "OPT":
            contract = ibi.Option(order[8], "20" + order[9].strftime("%y%m%d"), order[10], order[11], 'SMART', multiplier=100, currency='USD')
            ib.qualifyContracts(contract)
        else:
            raise ValueError ("Incorrect instrument type")

    # Disconnect from IB after placing the orders
    ib.disconnect()


def globalCancelOrders():
    ib.reqGlobalCancel()


def main():
    global ib, wb

    #controller = ibi.IBController('TWS', '969', 'paper',
    #    TWSUSERID='enter username here', TWSPASSWORD='enter password here')
    ib = ibi.IB()

    #Create an object for the caller Excel workbook and note the calling worksheet
    wb = xw.Book.caller()

    while True:
        # Create an dictionary initialized with account numbers and empty lists
        # These lists will be populated with positions for the account
        acccountPositionsDict = {}
        for accountNumber in getAccountNumberDict():
            acccountPositionsDict[accountNumber] = []

        # Start IB connection
        ib.connect('127.0.0.1', getIbConnectionPort(wb.sheets.active.name), clientId=1)

        # Get positions for all accounts and organize them into the dictionary of lists created above
        for acccountPosition in ib.positions():
            acccountPositionsDict[acccountPosition.account].append([acccountPosition.contract.localSymbol, float(acccountPosition.avgCost), float(acccountPosition.avgCost)*float(acccountPosition.position)])

        # Update the worksheets with positions
        for accountNumber in acccountPositionsDict:
            sheetName = getAccountNumberDict().get(accountNumber)

            #Set NVL and Cash values in Excel cells
            wb.sheets(sheetName).range("B1").value = accountValue(accountNumber, 'NetLiquidationByCurrency')
            wb.sheets(sheetName).range("D1").value = accountValue(accountNumber, 'CashBalance')
            acccountPositionsDict.get(accountNumber).sort()

            # Get the Excel cell address where "Portfolio" table begins
            cellAddress = wb.sheets(sheetName).tables[sheetName + 'PortfolioTable'].data_body_range(1, 1).address

            wb.sheets(sheetName).range(cellAddress).value = acccountPositionsDict.get(accountNumber)

        # Close IB connection
        ib.disconnect()
        
        ib.sleep(60)

if __name__ == "__main__":
    xw.Book("ib_insync_XL.xlsm").set_mock_caller()
    main()
