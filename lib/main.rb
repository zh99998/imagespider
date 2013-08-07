class Collection
  attr_reader :hash
  def initialize
  	@hash_lib = {}
  	@hash_sql = {}
  	@hash_js = {}
  	@replace = {}
  end
  def load_lib
 	  require 'win32ole'
  	  conn = WIN32OLE.new('ADODB.Connection')
      conn.open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "YGODAT.DAT" + ";Jet OLEDB:Database Password=paradisefox@sohu.com" )
      records = WIN32OLE.new('ADODB.Recordset')
      records.open("YGODATA", conn)
      records.MoveNext
      while !records.EOF
      	name = records.Fields.Item("JPCardName").value
      	name.gsub!("－","−")
      	name.delete!(" ")
      	id = records.Fields.Item("CardPass").value
      	@hash_lib[id.to_i] = name
      	records.MoveNext
      end
      return @hash_lib
  end
  def load_js
  	  require 'json'
  	  f = File.open("orenoturn.json","r")
  	  data = JSON.parse(f.read)
  	  f.close
  	  for h in data
  	  	id = h["orenoturn_id"]
  	  	name = h["orenoturn_name"]
  	  	name = name.split(/【.+?】/)[0]
  	  	name.delete!("　")
  	  	name.delete!(" ")
  	  	@hash_js[name] = id
  	  	if (name.index("·") != nil)
  	  		name.gsub!("·",".")
  	  		@hash_js[x2] = id
  	  	end
  	  	if (name.index("・") != nil)
  	  		x2 = name.gsub("・","·")
  	  		@hash_js[x2] = id 
  	  	end
  	  end
  	  return @hash_js
  end
  def load_sql
  	require 'sqlite3'
    db = SQLite3::Database.new("cards.cdb")
    ar = db.execute("select id, name from texts")
    for a in ar
    	id = a[0]
    	name = a[1]
  	  	name.delete!("　")
  	  	name.delete!(" ")
    	@hash_sql[id] = name
    end
    return @hash_sql
  end
  def merge()
  	load_lib
  	load_js
  	load_sql
  	ans = {}
  	libs = []
  	sqls = []
  	for n1 in @hash_lib.keys
  		x = @hash_js[@hash_lib[n1]]
  		if (x == nil) 
  		 libs.push @hash_lib[n1]
  		else ans[n1] = x
  		end
  	end
  	for n1 in @hash_sql.keys
  		x = @hash_js[@hash_sql[n1]]
  		if(x == nil) 
  			sqls.push @hash_sql[n1]
  		else ans[n1] = x
  		end
  	end

  	sub = @hash_js.keys - (@hash_lib.values - libs) - (@hash_sql.values - sqls)
	f = File.open("lib.txt","w")
  	str = "以下在数据库中的卡片未被索引到（合计#{libs.size}）：\n"
  	for t in libs
  		str += t + "\n"
  	end
  	f.write(str)
  	f.close()
	f = File.open("sql.txt","w")
  	str = "以下在SQL中的卡片未被索引到（合计#{sqls.size}）：\n"
  	for t in sqls
  		str += t + "\n"
  	end
  	f.write(str)
  	f.close()
  	f = File.open("json.txt", "w")
  	str = "以下在json上的卡片未被索引到（合计#{sub.size}）：\n"
  	for t in sub 
  		str += t + "\n"
  	end
  	f.write(str)
  	f.close()
  	return ans
  end
end

f = File.open("x.txt","w")
f.write("=============================================\n")
f.flush()
begin
	ans = Collection.new.merge
	f.write("合计 #{ans.keys.size} 个匹配\n")
	f.write(ans)
	f.write("=============================================\n")
rescue => ex
	f.write(ex.to_s)
	f.write(ex.backtrace.to_s)
end
f.close