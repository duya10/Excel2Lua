require "luacom"
require "utf8gbk"
require "prettyprint"

-- 编码转换
G2U = function(s) return utf8gbk(s, true) end
U2G = function(s) return utf8gbk(s) end

function write_file(filename, str)
    local f = assert(io.open(filename, 'w'))
    f:write(str)
    f:close()
end

SEPARATOR = "\\"

function make_dir(directory)
    assert(type(directory) == "string")
    local path = nil
    if directory:sub(2, 2) == ":" then
        path = directory:sub(1, 2)
        directory = directory:sub(4)
    else
        if directory:match("^/") then
            path = ""
        end
    end
    print(directory)
    for d in directory:gmatch("([^" .. SEPARATOR .. "]+)" .. SEPARATOR .. "*") do
        path = path and path .. SEPARATOR .. d or d
        local mode = lfs.attributes(path, "mode")
        if not mode then
            local ok, err = lfs.mkdir(path)
            if not ok then
                return false, err
            end
        elseif mode ~= "directory" then
            return false, path .. " is not a directory"
        end
    end
    return true
end

function find_files(dir_path, finded_file_list, extension)
    local finded_file_list = finded_file_list or {}
    for entry in lfs.dir(dir_path) do
        if entry ~= "." and entry ~= ".." then
            local file_path = dir_path .. SEPARATOR .. entry
            local file_attr = lfs.attributes(file_path)
            if file_attr.mode == "directory" then
                find_files(file_path, finded_file_list, extension)
            elseif file_attr.mode == "file" and string.find(entry, "^~") == nil and string.find(entry, extension) then
                local file_info = {
                    file_path           = file_path,
                    entry               = entry,
                    dir_path            = dir_path,
                    last_modification   = file_attr.modification,
                    last_size           = file_attr.size,
                }
                finded_file_list[file_path] = file_info
            end
        end
    end
end

function class(classname, super)
    local superType = type(super)
    local cls

    if superType ~= "function" and superType ~= "table" then
        superType = nil
        super = nil
    end

    if super then
        cls = {}
        setmetatable(cls, {__index = super})
        cls.super = super
    else
        cls = {ctor = function() end}
    end

    cls.__cname = classname
    cls.__index = cls

    function cls.new(...)
        local obj_data = {}
        local cls_name = cls.__cname

        local instance = setmetatable(obj_data, cls)
        instance.class = cls
        instance:ctor(...)
        return instance
    end
    return cls
end