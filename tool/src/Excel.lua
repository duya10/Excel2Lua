local Sheet = require "Sheet"
excel_application = excel_application or luacom.CreateObject("Excel.Application")

local Excel = {}

function Excel.new()
    local o = {}
    setmetatable(o, Excel)
    Excel.__index = Excel
    return o
end

function Excel:open(path)
    assert( self:isExist(path), path )

    local excel = excel_application
    excel.Visible = 0
    assert(excel, path)
    self.excel = excel

    excel.Application.DisplayAlerts   = 0
    excel.Application.ScreenUpdataing = 0
    local book = excel.WorkBooks:open(G2U(path), nil, 0)

    book.Saved = false
    self.book = book
    return true
end

function Excel:close()
    if self.book then
        self.book:Close()
    end
    if self.excel then
        self.excel.Application:Quit()
    end
end

function Excel:selectSheet(sheet)
    self.book.Sheets(sheet):Select()
    return Sheet.new(self.excel.ActiveSheet, self)
end

function Excel:isExist(path)
    local t = io.open(path, 'r')
    if not t then
        return false
    else
        t:close()
        return true
    end
end

return Excel