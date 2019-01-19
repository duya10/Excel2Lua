--[[
    序列化和反序列化 lua objects or table
]]
local LuaText = class("LuaText")

function LuaText:ctor()

end

function LuaText:_serialize(value, indent)
    if not indent then indent = 4 end
    local maxLines = 50000 -- for sanity
    local maxChars = maxLines * 200 -- 10MB

    local s = ""
    local strArray = {} -- 字符串数组
    local lines = 1
    local tableLabel = {}
    local nTables = 0

    local function addNewLine(i)
        if #s >= maxChars or lines >= maxLines then return true end
        if indent > 0 then
            table.insert(strArray, "\n" .. string.rep(" ", i))
            lines = lines + 1
        end
        return false
    end

    local function appendOne(x, i)
        if type(x) == "string" then
            local tmpStr = string.format("%q", x):gsub("\n", "n")
            table.insert(strArray, tmpStr)

        elseif type(x) ~= "table" then
            table.insert(strArray, tostring(x))

        elseif type(getmetatable(x)) == "string" then
            table.insert(strArray, tostring(x))

        else
            if tableLabel[x] then
                return false
            end

            local isEmpty = true
            for k, v in pairs(x) do isEmpty = false; break end
            if isEmpty then
                table.insert(strArray, "{}")
                return false
            end

            nTables = nTables + 1
            local label = "table: " .. nTables
            tableLabel[x] = label

            table.insert(strArray, "{")

            local first = true

            local sort_table = {}
            for k, v in pairs(x) do
                table.insert(sort_table, {key = k})
            end
            table.sort(sort_table, function(a, b) return a.key < b.key end)

            for _, val in ipairs(sort_table) do
                local k = val.key
                local v = x[k]
                if first then
                    first = false
                else
                    table.insert(strArray, ", ")
                end
                if addNewLine(i + indent) then return true end
                if type(k) == "string" and k:match("^[_%a][_%w]*$") then
                    table.insert(strArray, k)
                else
                    table.insert(strArray, "[")
                    if appendOne(k, i + indent) then return true end
                    table.insert(strArray, "]")
                end
                table.insert(strArray, " = ")
                if appendOne(v, i + indent) then return true end
            end

            table.insert(strArray, "}")
        end

        return false
    end

    local v = appendOne(value, 0)

    s = table.concat(strArray)
    return s
end

function LuaText:serialize(t)
    local s = self:_serialize(t, 4)
    return s
end

function LuaText:_deserialize(s, name)
    local func, err = loadstring(s, name)
    if func then
        local result = {}
        result[1], result[2] = func()
        if result[2] ~= nil then
            error("Custom lua deserializer only support a value")
        else
            return result[1]
        end
    else
        local x = err:find(name)
        if x then
            err = err:sub(x)
        elseif err:len() > 77 then
            err = err:sub(-77)
        end
        return nil, err
    end
end

function LuaText:deserialize(s, name, obj)
    if not str:is(name) then
        if o then
            name = str:to(o)
        else
            name = "unknown chunk"
        end
    end
    local t, err = self:_deserialize(s)
    if t then
        if not o then
            return t
        else
            for k, v in pairs(t) do
                o[k] = v
            end
            return o
        end
    else
        return nil, err
    end
end

return LuaText