local Sheet = {}

function Sheet.new(ptr, excel)
    local o = {sheet = ptr, owner = excel}
    setmetatable(o, Sheet)
    Sheet.__index = Sheet
    return o
end

function Sheet:getCell(column, row)
    return self.sheet.Cells(column, row).Value2
end

function Sheet:Range(range)
    return self.sheet:Range(range).Value2
end

function Sheet:getUseRange(range)
    return self.sheet.Usedrange.Rows.count, self.sheet.Usedrange.columns.count
end

-- 进制转换
local function dec2X(dec, x)
    local new_number = {}

    local function f(a)
        assert(a >= 1)
        local mod = dec % math.pow(x, a)
        local last_mod = (a == 1) and 0 or assert(new_number[a - 1])
        new_number[a] = (mod - last_mod) / math.pow(x, a - 1)
        -- 取整数部分
        new_number[a] = math.modf(new_number[a])
        return mod ~= dec
    end

    local i = 1
    while f(i) do
        i = i + 1
    end

    return new_number
end

-- 数组倒序
local function orderByDesc(input)
    local output = {}
    local count = #input
    while count > 0 do
        table.insert(output, input[count])
        count = count - 1
    end
    return output
end

function Sheet:getColumnString(num)
    local number_tbl = dec2X(num - 1, 26)
    number_tbl[1] = number_tbl[1] + 1

    for i, v in ipairs(number_tbl) do
        number_tbl[i] = string.char(string.byte('A') + v - 1)
    end
    number_tbl = orderByDesc(number_tbl)

    local s = table.concat(number_tbl)
    return s
end

assert(Sheet.getColumnString(nil, 1) == 'A')
assert(Sheet.getColumnString(nil, 27) == 'AA')
assert(Sheet.getColumnString(nil, 256) == 'IV')
assert(Sheet.getColumnString(nil, 26) == 'Z')

return Sheet