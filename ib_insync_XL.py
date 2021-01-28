import ib_insync as ibi
import xlwings as xw
import asyncio

def accountValue(accountNumber, tag: str) -> str:
    """Return account value for the given tag."""
    return next((
        v.value for v in ib.accountValues(accountNumber)
        if v.tag == tag and v.currency == 'BASE'), '')


def updateAccountAndPortfolio():

    # Create an empty dictionary with each account number and empty lists that will be populaqted with positions
    acccountPositionsDict = {}
    for accountNumber in accountNumberToSheet:
        acccountPositionsDict[accountNumber] = []

    # Get all positions for all accounts and organize them into a dictionary of lists
    for acccountPosition in ib.positions():
        #ib.reqPnLSingle(account=portfolioItem.account, conId=portfolioItem.contract.conId, modelCode='')
        acccountPositionsDict[acccountPosition.account].append([acccountPosition.contract.localSymbol, acccountPosition.position, float(acccountPosition.avgCost)*float(acccountPosition.position)])#, ib.PnLSingle(account=acccountPosition.account, conId=acccountPosition.contract.conId, modelCode='').get("unrealizedPnL"))#, acccountPosition.unrealizedPNL])
    
    # Update the worksheets with positions
    for accountNumber in acccountPositionsDict:
        sheetName = accountNumberToSheet.get(accountNumber)

        #Set NVL and Cash values in Excel cells
        wb.sheets(sheetName).range("B1").value = accountValue(accountNumber, 'NetLiquidationByCurrency')
        wb.sheets(sheetName).range("D1").value = accountValue(accountNumber, 'CashBalance')
        acccountPositionsDict.get(accountNumber).sort()

        # Get the Excel cell address where "Positions to Close" begin
        temp_macro = wb.macro("findTextOccurence")
        cellAddress = temp_macro(sheetName, 'Portfolio')

        #wb.sheets(sheetName).range(range(cellAddress), range(cellAddress).offset(0, 3).end('down')).clear_contents()
        wb.sheets(sheetName).range(cellAddress).options(ndim=2).value = acccountPositionsDict.get(acccountPosition.account)


def closePositions():
    wb = xw.Book.caller()
    ib = ibi.IB()
    closePositionsWorksheet = wb.sheets.active.name

    # Get the Excel cell address where "Positions to Close" begin
    temp_macro = wb.macro("findTextOccurence")
    cellAddress = temp_macro(closePositionsWorksheet, 'Positions to Close')

    # Start a new IB connection
    ib.connect('127.0.0.1', getIbConnectionPort(closePositionsWorksheet), clientId=2)

    positionsToClose = []

    # Retrieve positions to be closed
    if wb.sheets(closePositionsWorksheet).range(cellAddress).value.strip() == "":
        raise ValueError ("No Positions to Close")
    else:
        positionsToClose = wb.sheets(closePositionsWorksheet).range(cellAddress).options(expand='table').value
    
    # Within currently held positions, find the positions to be closed, and place closing orders
    for acccountPosition in ib.positions():
        for position in positionsToClose:
            if acccountPosition.contract.localSymbol.replace(" ", "") == position[0].replace(" ", ""):
                # if current position for a contract is LONG, then set closing action to "SELL"
                # otherwise set closing action to "BUY"
                closingAction = "SELL" if acccountPosition.position > 0 else "BUY"

                # Prepare closing order
                if position[1].upper() == "LMT":
                    # If limit price is not set, then raise error
                    if ibi.util.isNan(position[0]) or position[0] < 0:
                        ib.disconnect()
                        raise ValueError ("Incorrect Limit Price for " + position[0])
                    closingOrder = ibi.LimitOrder(closingAction, acccountPosition.position, position[2])
                elif position[1].upper() == "MKT":
                    closingOrder = ibi.LimitOrder(closingAction, acccountPosition.position)
                else:
                    raise ValueError ("Incorrect Order Type for " + position[0])

                ib.placeOrder(acccountPosition.contract, closingOrder)

    # Disconnect from IB after placing the orders
    ib.disconnect()


def getIbConnectionPort(xlSheetName):
    # return the socket port number to use for connecting to TWS
    if xlSheetName == "Paper":
        return 7497
    else:
        return 7496


def readAndPlaceXLOrders():
    accountNumberToSheet = {"Add account number here": "Brokerage"}

    wb = xw.Book.caller()
    ib = ibi.IB()
    ordersCallingWorksheet = wb.sheets.active.name

    # Get the Excel cell address where "Positions to Close" begin
    temp_macro = wb.macro("findTextOccurence")
    cellAddress = temp_macro(ordersCallingWorksheet, 'Order List')

    # Start a new IB connection
    ib.connect('127.0.0.1', getIbConnectionPort(ordersCallingWorksheet), clientId=2)

    ordersFromXL = []

    # Retrieve orders to be placed from Excel
    if wb.sheets(ordersCallingWorksheet).range(cellAddress).value.strip() == "":
        raise ValueError ("No Orders to Submit")
    else:
        ordersFromXL = wb.sheets(ordersCallingWorksheet).range(cellAddress).options(expand='table').value

    # Place orders one at a time
    for order in ordersFromXL:
        if order[2].upper() == "LMT":
            entryOrder = ibi.LimitOrder(order[1], int(order[5]), order[4])
        elif order[2].upper() == "MKT":
            entryOrder = ibi.MarketOrder(order[1], int(order[5]))
        else:
            raise ValueError ("Incorrect Order Type in " + order[0])

        if order[7].upper() == "STK":
            contract = ibi.Stock(order[7], 'SMART', 'USD', primaryExchange='NYSE')
        elif order[7].upper() == "OPT":
            contract = ibi.Option(order[7], "20" + order[8].strftime("%y%m%d"), order[9], order[10], 'SMART', multiplier=100, currency='USD')
            ib.qualifyContracts(contract)
        else:
            raise ValueError ("Incorrect instrument type")
        ib.placeOrder(contract, entryOrder)

    # Disconnect from IB after placing the orders
    ib.disconnect()


def globalCancelOrders():
    ib.reqGlobalCancel()


def main():
    global ib, wb, accountNumberToSheet
    accountNumberToSheet = {"U4529195":"Brokerage"}

    #controller = ibi.IBController('TWS', '969', 'paper',
    #    TWSUSERID='enter username here', TWSPASSWORD='enter password here')
    ib = ibi.IB()
    #ib.sleep(30)

    #Create an object for the caller Excel workbook and note the calling worksheet
    wb = xw.Book.caller()

    # Start IB api connection
    ib.connect('127.0.0.1', getIbConnectionPort(wb.sheets.active.name), clientId=1)

    loop = asyncio.get_event_loop()

    while True:
        updateAccountAndPortfolio()
        ib.sleep(30)

    try:
        loop.run_forever()
    except (KeyboardInterrupt, SystemExit):
        pass
    ib.disconnect()
    ib.sleep(0.1)

if __name__ == "__main__":
    xw.Book("ib_insync_XL.xlsm").set_mock_caller()
    main()
