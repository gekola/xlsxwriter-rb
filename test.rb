require 'xlsxwriter'

wb = XlsxWriter::Workbook.new('./test.xlsx', constant_memory: true)
wb
  .add_format(:numf, num_format_index: 0x04, font_color: XlsxWriter::Format::COLOR_GREEN)
  .add_format(:date, num_format_index: 0x0e)
  .add_format(:datetime, num_format_index: 0x16)

ws = wb.add_worksheet('ttt')
ws2 = wb.add_worksheet('zzz')
1_000_000.times do |i|
  # ws.write_number(i,0,rand*10_000, :numf)
  # ws.write_string(i,1,('testzzz%0d' % (rand * 1000)),nil)
  ws.add_row([rand * 10_000, 'testzzz%0d' % (rand * 1000)], style: [:numf])
end

12_000.times do |i|
  ws2.merge_range(i, 'P', i, 'R', '', nil)
  ws2.add_row([
               '%d' % (rand * 100),
               rand,
               '=B1',
               Date.new(2010, 3, 5),
               'http://example.com/',
               rand > 0.5,
               '',
               rand * 10,
               Time.now,
               rand < 0.5,
               '',
               '=A1',
               'http://example.com/2',
               'zzzz',
               nil,
               Object.new
              ],
              types: [:string, :number, :formula, :datetime, :url, :boolean, :blank],
              style: [nil, nil, nil, :date, nil, nil, nil, nil, :datetime])
end

ws2.insert_image(13, 'C', 'ext/xlsxwriter/libxlsxwriter/test/functional/src/images/red.png', {scale: 2.0, x_offset: 10, y_offset: 3})
