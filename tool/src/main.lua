-- 获取路径参数
local args = {}
for word in string.gmatch(..., "%S+") do
    table.insert(args, word)
end
local DATA_PATH, XLSX_PATH, LUA_PATH = args[1], args[2], args[3]

package.cpath = package.cpath .. ";../bin/?.dll"
-- package.path  = package.path .. ";../../"

require "util"
local lfs     = require "lfs"
local Excel   = require "Excel"
local LuaText = require "LuaText"

local luatextObj = LuaText.new()

local function trans_to_type(all_range, max_read_rows, row, col)
    local index = math.ceil(row / max_read_rows)
    local range = all_range[index]
    local real_row = row - (index - 1) * max_read_rows
    local cell_data = range[real_row][col]
    if not cell_data then
        return
    end
    local val = cell_data
    val = tostring(val)
    return val, cell_data
end

local NODE_ARRAY = 1
local NODE_MAP   = 2
local NODE_ROOT  = 3
local function excel_to_lua(path, export_file_path, exclude_key)
    exclude_key = exclude_key or ""
    local docObj = Excel.new()
    docObj:open(path)

    local sheetObj = docObj:selectSheet(1)

    local rows, columns = sheetObj:getUseRange()
    print(U2G(string.format("行 %d，列 %d", rows, columns)))
    print("export " .. path)

    -- 获取表头
    local table_head = {}
    for c = 1, columns do
        local cn_name = sheetObj:getCell(1, c)
        local en_name = sheetObj:getCell(2, c)
        local node    = sheetObj:getCell(3, c)
        if en_name == nil then
            break
        end
        assert(en_name:match("%s+") == nil, U2G(string.format("表[%s]的[%s]字段存在空白符号", path, en_name)))

        local node_type
        if node == "root" then
            node_type = NODE_ROOT
        elseif node == "map" then
            node_type = NODE_MAP
        elseif node == "array" then
            node_type = NODE_ARRAY
        end
        if c == 1 and node_type == nil then
            node_type = NODE_ROOT
        end
        table_head[c] = { cn_name = cn_name, en_name = en_name, node_type = node_type }
    end
    print(table_head)
    columns = #table_head

    local beg_column = sheetObj:getColumnString(1)
    local end_column = sheetObj:getColumnString(columns)

    print("beg_column ",beg_column, ", end_column ", end_column)

    local all_range = {}
    local MAX_READ_CELLS = 3000
    local max_read_rows = math.floor(MAX_READ_CELLS / columns)
    local read_range_times = math.ceil(rows / max_read_rows)
    for index = 1, read_range_times do
        local begin_row = (index - 1) * max_read_rows + 1
        local end_row   = index * max_read_rows
        if end_row > rows then
            end_row = rows
        end
        local range_str = string.format("%s%s:%s%s", beg_column, begin_row, end_column, end_row)
        local range_tbl = sheetObj:Range(range_str)
        print(range_tbl)
        all_range[index] = range_tbl
    end

    -- 读取数据
    local BEG_ROW = 4
    local table_data = {}
    local all_rows = {}
    for r = BEG_ROW, rows do
        local key 
        local root_table = table_data
        local is_empty_row = true
        local row_info = {}
        all_rows[r] = row_info
        for c = 1, columns do
            local en_name   = table_head[c].en_name
            local node_type = table_head[c].node_type
            local val, real_val = trans_to_type(all_range, max_read_rows, r, c)
            if not real_val then
                if node_type == NODE_ROOT then
                    local last_row_info = all_rows[r - 1]
                    if last_row_info == nil then
                        break
                    end
                    val = last_row_info[en_name]
                end
            else
                is_empty_row = false
            end
            if node_type == NODE_ROOT then
                if root_table[val] == nil then
                    root_table[val] = {}
                end
                root_table = root_table[val]
            elseif node_type == NODE_MAP then
                key = val
            elseif node_type == NODE_ARRAY then
                key = #root_table + 1
            end
            if string.find(en_name, exclude_key) == nil and val then
                row_info[en_name] = val
            end
        end
        if is_empty_row then
            break
        end
        if key then
            root_table[key] = row_info
        end
    end

    docObj:close()

    write_file(export_file_path, "return " .. luatextObj:_serialize(table_data))
    return table_data
end

function main()
    local record_path = "../bin/record_files.lo"
    local all_records = {}
    local record_files = io.open(record_path, 'rb')
    if record_files then
        local content = record_files:read('*a')
        all_records = loadstring("return " .. content)() or {}
    end

    local remove_lua_dir = {}
    local modify_lua_dir = {}

    local all_xlsx, all_lua = {}, {}
    find_files(XLSX_PATH, all_xlsx, "%.xlsx$")
    for file_path, file_info in pairs(all_xlsx) do
        local dir_path          = file_info.dir_path
        local file_name         = file_info.file_name

        local export_file_path = string.gsub(file_path, XLSX_PATH, LUA_PATH)
        export_file_path = string.gsub(export_file_path, ".xlsx$", ".lua")

        local record_info = all_records[file_path]
        local is_export_file_exist = false
        local export_file = io.open(export_file_path, 'rb')
        if export_file then
            -- is_export_file_exist = true
            export_file:close()
        end
        if record_info == nil
            or record_info.last_modification ~= file_info.last_modification
            or record_info.last_size ~= file_info.last_size
            or not is_export_file_exist then

            local lua_dir = string.gsub(dir_path, XLSX_PATH, LUA_PATH)
            make_dir(lua_dir)
            all_lua[export_file_path] = excel_to_lua(file_path, export_file_path)
            modify_lua_dir[lua_dir] = 1

        end
        all_records[file_path] = nil
    end

    for file_path, file_info in pairs(all_records) do
        local export_file_path = string.gsub(file_path, XLSX_PATH, LUA_PATH)
        export_file_path = string.gsub(export_file_path, ".xlsx$", ".lua")
        os.remove(export_file_path)
        local export_dir_path = string.gsub(file_info.dir_path, XLSX_PATH, LUA_PATH)
        remove_lua_dir[export_file_path] = 1
    end

    write_file(record_path, luatextObj:_serialize(all_xlsx))
    excel_application.Application:Quit()

    print(U2G("导表完成"))
    -- os.execute("TortoiseProc.exe /command:commit /path:../../ / closeonend:3 /logmsg:" .. U2G("服务器导表提交"))
end

local ok, msg = pcall(main)
if not ok then
    print(msg)
    print(U2G("导表失败\n"))
    excel_application.Application:Quit()
end