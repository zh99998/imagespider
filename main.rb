#encoding: UTF-8

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

  attr_reader :hash

  def initialize
    @hash_lib = {}
    @hash_sql = {}
    @hash_js  = {}
    @replace  = {}
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
    f.close
    for h in data
      next if h['orenoturn_image_basename'][0,11] == 'noimage.jpg'
      id             = h["orenoturn_id"]
      name           = h["orenoturn_name"]
      name           = name.split(/【.+?】/)[0]
      name           = characters(name)
      @hash_js[name] = id
    end
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
    ans  = {}
    libs = []
    sqls = []
    for n1 in @hash_lib.keys
      x = @hash_js[@hash_lib[n1]]
      if (x == nil)
        libs.push @hash_lib[n1]
      else
        ans[n1] = x
      end
    end
    for n1 in @hash_sql.keys
      x = @hash_js[@hash_sql[n1]]
      if (x == nil)
        sqls.push @hash_sql[n1]
      else
        ans[n1] = x
      end
    end

    sub = @hash_js.keys - (@hash_lib.values - libs) - (@hash_sql.values - sqls)
    f   = File.open("lib.txt", "w")
    str = "以下在数据库中的卡片未被索引到（合计#{libs.size}）：\n"
    for t in libs
      str += t + "\n"
    end
    f.write(str)
    f.close()
    f   = File.open("sql.txt", "w")
    str = "以下在SQL中的卡片未被索引到（合计#{sqls.size}）：\n"
    for t in sqls
      str += t + "\n"
    end
    f.write(str)
    f.close()
    f   = File.open("json.txt", "w")
    str = "以下在json上的卡片未被索引到（合计#{sub.size}）：\n"
    for t in sub
      str += t + "\n"
    end
    f.write(str)
    f.close()
    return ans
  end
end

ans = Collection.new.merge
open('result.json', 'w:UTF-8'){|f|f.write ans.to_json}