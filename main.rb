#encoding: UTF-8

$log = File.open("log.log","w:UTF-8")

class Collection

  Alleles = {
      'Ａ'  => 'A',
      'Ｂ'  => 'B',
      'Ｃ'  => 'C',
      'Ｄ'  => 'D',
      'Ｅ'  => 'E',
      'Ｆ'  => 'F',
      'Ｇ'  => 'G',
      'Ｈ'  => 'H',
      'Ｉ'  => 'I',
      'Ｊ'  => 'J',
      'Ｋ'  => 'K',
      'Ｌ'  => 'L',
      'Ｍ'  => 'M',
      'Ｎ'  => 'N',
      'Ｏ'  => 'O',
      'Ｐ'  => 'P',
      'Ｑ'  => 'Q',
      'Ｒ'  => 'R',
      'Ｓ'  => 'S',
      'Ｔ'  => 'T',
      'Ｕ'  => 'U',
      'Ｖ'  => 'V',
      'Ｗ'  => 'W',
      'Ｘ'  => 'X',
      'Ｙ'  => 'Y',
      'Ｚ'  => 'Z',
      'ａ'  => 'a',
      'ｂ'  => 'b',
      'ｃ'  => 'c',
      'ｄ'  => 'd',
      'ｅ'  => 'e',
      'ｆ'  => 'f',
      'ｇ'  => 'g',
      'ｈ'  => 'h',
      'ｉ'  => 'i',
      'ｊ'  => 'j',
      'ｋ'  => 'k',
      'ｌ'  => 'l',
      'ｍ'  => 'm',
      'ｎ'  => 'n',
      'ｏ'  => 'o',
      'ｐ'  => 'p',
      'ｑ'  => 'q',
      'ｒ'  => 'r',
      'ｓ'  => 's',
      'ｔ'  => 't',
      'ｕ'  => 'u',
      'ｖ'  => 'v',
      'ｗ'  => 'w',
      'ｘ'  => 'x',
      'ｙ'  => 'y',
      'ｚ'  => 'z',
      '０'  => '0',
      '１'  => '1',
      '２'  => '2',
      '３'  => '3',
      '４'  => '4',
      '５'  => '5',
      '６'  => '6',
      '７'  => '7',
      '８'  => '8',
      '９'  => '9',
      '－'  => '-',
      '−'  => '-',
      '．'  => '',
      '·'  => '',
      '・'  => '',
      '.'  => '',
      '／'  => '/',
      '　'  => '',
      ' '  => '',

      #single fix
      '龍'  => '竜', #60107471
      '期'  => '後', #60107367
      'R】' => '', #59444207
      '罠'  => '網', #59119904
      '帚'  => '箒',
      'ぜ'  => 'ゼ',
      '-'  => '',
      'キ'  => '璣',
      'セン' => '璇',
      'コウ' => '罡',
      'レ' => '',
      '―' => 'ー',
      '’' => '',

      'No32破滅のアシッドゴーム' => 'No30破滅のアシッドゴーム', # FUUUUUUUUUUUCK
      'エンジェル07' => 'エンジェルO7',
      'BARRELDRAGON' => 'リボルバードラゴン',
      'SLIFERTHESKYDRAGON' => 'オシリスの天空竜',
      'OBELISKTHETORMENTOR' => 'オベリスクの巨神兵',
      'THEWINGEDDRAGONOFRA' => 'ラーの翼神竜'

  }

  Score = 
  {
  	"【N】"   => 10,
  	"【NR】"  => 10,
  	"【SR】"  => 8,
  	"【R】"   => 7,
  	"【NPR】" => 5,
  	"【GR】"  => 4,
  	"【UR】"  => 3,
  	"【UTR】" => 2,
  	"【SC】"  => 1,
  	"" => 0,
  	nil => 0
  }

  attr_reader :hash

  def initialize
    @hash_lib = {}
    @hash_sql = {}
    @hash_js  = {}
    @replace  = {}
    @ans      = {}
  end

  def load_lib
    require 'win32ole'
    conn = WIN32OLE.new('ADODB.Connection')
    conn.open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=YGODAT.DAT;Jet OLEDB:Database Password=paradisefox@sohu.com")
    records = WIN32OLE.new('ADODB.Recordset')
    records.open("YGODATA", conn)
    records.MoveNext
    while !records.EOF
      name               = records.Fields.Item("JPCardName").value
      name               = characters(name)
      id                 = records.Fields.Item("CardPass").value
      @hash_lib[id.to_i] = name
      records.MoveNext
    end
    return @hash_lib
  end

  def load_js
    require 'json'
    f    = File.open("orenoturn.json", "r:UTF-8")
    data = JSON.parse(f.read)
    marks = {}    # 罕贵得分
    f.close
    for h in data
      next if h['orenoturn_image_basename'][0,11] == 'noimage.jpg'
      id             = h["orenoturn_id"]
      name           = h["orenoturn_name"]
      exps           = /【.+?】/.match(name)
      exps           = exps.to_s
      name           = name.split(/【.+?】/)[0]
      name           = characters(name)

      #$log.write("#{name} 的罕贵为 #{exps}，得分为#{Score[exps]}\n" )
      mark           = Score[exps]   # 得分判定
      mark           = 0 if mark == nil
      if marks[name] != nil          # 若此项业已存在
      	if mark > marks[name]          # 得分高者替换
      		$log.write("进行了罕贵替换： #{name} 被替换成了罕贵：#{exps}\n")
      		marks[name]    = mark
      		@hash_js[name] = id
      	else                         # 得分低者忽略
      	end
      else
     	 @hash_js[name] = id
     	 marks[name] = mark
      end
    end
    #for key in marks.keys
    #	$log.write("#{key}最终的罕贵为#{marks[key]}\n")
    #end
    return @hash_js
  end

  def load_sql
    require 'sqlite3'
    db = SQLite3::Database.new("cards.cdb")
    ar = db.execute("select id, name from texts")
    for a in ar
      id            = a[0]
      name          = a[1]
      name          = characters(name)
      @hash_sql[id] = name
    end
    return @hash_sql
  end

  def characters(str)
    str.encode!("UTF-8")
    Alleles.each_pair { |key, value| str.gsub!(key, value) }
    str
  end

  def merge()
    load_lib
    load_js
    load_sql
    @ans  = {}
    libs = []
    sqls = []
    for n1 in @hash_lib.keys
      x = @hash_js[@hash_lib[n1]]
      if (x == nil)
        libs.push @hash_lib[n1]
      else
        @ans[n1] = x
      end
    end
    for n1 in @hash_sql.keys
      x = @hash_js[@hash_sql[n1]]
      if (x == nil)
        sqls.push @hash_sql[n1]
      else
        @ans[n1] = x
      end
    end

    sub = @hash_js.keys - (@hash_lib.values - libs) - (@hash_sql.values - sqls)
    str = "以下在数据库中的卡片未被索引到（合计#{libs.size}）：\n"
    for t in libs
      str += t + "\n"
    end
    $log.write(str)
    str = "以下在SQL中的卡片未被索引到（合计#{sqls.size}）：\n"
    for t in sqls
      str += t + "\n"
    end
    $log.write(str)
    str = "以下在json上的卡片未被索引到（合计#{sub.size}）：\n"
    for t in sub
      str += t + "\n"
    end
    $log.write(str)
    return @ans
  end

  def check_lib()
    require 'win32ole'
    conn = WIN32OLE.new('ADODB.Connection')
    conn.open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=YGODAT.DAT;Jet OLEDB:Database Password=paradisefox@sohu.com")
    records = WIN32OLE.new('ADODB.Recordset')
    records.open("YGODATA", conn)
    records.MoveNext
    id2pas = {}
    id2name = {}
    while !records.EOF
      id                 = records.Fields.Item("CardID").value.to_i
      pas                = records.Fields.Item("CardPass").value.to_i
      name               = records.Fields.Item("SCCardName").value.to_s
      id2pas[id] = pas
      id2name[id] = name
      records.MoveNext
    end
    id2pic = {}
    noimages = []
    for key in id2pas.keys
      pic = @ans[id2pas[key]]
      if pic != nil then id2pic[key] = pic   # 不能一行的 Ruby 逼我用了 then
      else noimages.push key
      end
    end
    str = ""                                  # 我遇到了一个奇怪的编码问题 所以 str 写到后面去了
    for x in noimages
    	str += id2name[x] + "\n"
    end
    $log.write("以下在数据库中的卡片未被链接至图像（合计#{noimages.size}）：\n"); 
    $log.write(str)
    return id2pic
  end
end

collection = Collection.new
ans = collection.merge
open('result.json', 'w:UTF-8'){|f|f.write ans.to_json}
lib = collection.check_lib
open('lib.json', 'w:UTF-8') {|f|f.write lib.to_json}
$log.close
