require 'bundler'
Bundler.require
require 'pp'

workbook = RubyXL::Parser.parse('sample/xlsx/gooddata.xlsx')
worksheet = workbook[0]

md = File.open('sample/md/gooddata.md', 'w')

i = 0
h2flag = false
tableflag = false
worksheet.extract_data.each do |row|

  # 見出し1の処理
  if i == 0
    row.unshift('# ')
    row = row.first(2).join
  end

  # 見出し2の処理
  if row.nil?
    h2flag = true
    md.puts "\n"
    next
  end
  if h2flag
    row = row.unshift('## ').first(2).join
  end

  # リストの処理
  if row[0].nil?
    row = row.compact.unshift('- ').join
    md.puts row
    next
  end

  # 表の処理
  if row.is_a?(Array)
    if tableflag
      row = row.join(' | ')
      row = '| ' + row + ' |'
      md.puts row
      next
    end
    if row.count > 1
      cnt = row.size
      row = row.join(' | ')
      row = '| ' + row + ' |'
      md.puts row
      tableflag = true

      thline = []
      cnt.times do
        thline << '---'
      end
      thline = thline.join(' | ')
      md.puts '| ' + thline + ' |'
      next
    end
  end

  md.puts row
  md.puts "\n"
  i += 1
  h2flag = false
  tableflag = false
end
