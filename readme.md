# Excel to Lua
The first row will be ignored, this row is used for designer's description.

The second row is used for the generated Lua's field.

Support field types:
1. number
2. string
3. array
4. bool

If you want to ignore the field, use the hyphen character(-).

Example:
```
唯一ID	名称	忽略列	数值	数组	是否
Id:number	Name:string	-	Float:number	Array:array	Bool:bool
1	中华	-	1.2444	{1,2,3}	TRUE
2	消息	-	3	{1,2,4}	TRUE
3	心宿二	-	4	{1,2,5}	FALSE
4	短短的	-	5	{1,2,6}	TRUE
5	啊啊啊	-	5	{1,2,7}	TRUE

```

```
go run xls2lua.go
```

Show the results:
```
lua test.lua
```
