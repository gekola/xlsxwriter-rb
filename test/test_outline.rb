# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestOutline < XlsxWriterTestCase
  test 'outline01' do |wb, t|
    t.ignore_files << 'xl/calcChain.xml' << '[Content_Types].xml' << 'xl/_rels/workbook.xml.rels'

    wb.add_format(:bold, bold: true)
    ws = wb.add_worksheet 'Outlined Rows'
    ws.set_column('A', 'A', width: 20)
    (1..10).each do |i|
      ws.set_row i, level: ((i % 5).zero? ? 1 : 2)
    end

    write_common_data(ws)
  end

  test 'outline02' do |wb, t|
    t.ignore_files << 'xl/calcChain.xml' << '[Content_Types].xml' << 'xl/_rels/workbook.xml.rels'

    wb.add_format(:bold, bold: true)
    ws = wb.add_worksheet 'Collapsed Rows'
    (1..10).each do |i|
      ws.set_row i, level: ((i % 5).zero? ? 1 : 2), hide: true
    end
    ws.set_row 11, collapse: true

    ws.set_column('A', 'A', width: 20)
    ws.set_selection('A14:A14')

    write_common_data(ws)
  end

  test 'outline03' do |wb, t|
    t.ignore_files << 'xl/calcChain.xml' << '[Content_Types].xml' << 'xl/_rels/workbook.xml.rels'

    wb.add_format(:bold, bold: true)
    ws = wb.add_worksheet 'Outline Columns'

    %w[Month Jan Feb Mar Apr May Jun Total].each_with_index { |s, i| ws.write_string 0, i, s }
    [
      ['North', [50, 20, 15, 25, 65, 80]],
      ['South', [10, 20, 30, 50, 50, 50]],
      ['East',  [45, 75, 50, 15, 75, 100]],
      ['West',  [15, 15, 55, 35, 20, 50]]
    ].each_with_index do |(d, vals), i|
      ws.write_string(i + 1, 0, d)
      vals.each_with_index { |v, j| ws.write_number(i + 1, j + 1, v) }
      ws.write_formula_num(i + 1, vals.size + 1, "=SUM(B#{i + 2}:G#{i + 2})", vals.reduce(:+))
    end
    ws.write_formula_num('H6', '=SUM(H2:H5)', 1015, :bold)
    ws.set_row 0, format: :bold
    ws.set_column 'A', 'A', width: 10, format: :bold
    ws.set_column 'H', 'H', width: 10
    ws.set_column 'B', 'G', width: 6, level: 1
  end

  test 'outline04' do |wb, t|
    t.ignore_files << 'xl/calcChain.xml' << '[Content_Types].xml' << 'xl/_rels/workbook.xml.rels'

    ws = wb.add_worksheet 'Outline levels'
    13.times { |i| ws.write_string i, 0, "Level #{i > 6 ? 13 - i : i + 1}" }
    13.times { |i| ws.set_row i, level: (i > 6 ? 13 - i : i + 1) }
  end

  test 'outline05' do |wb, t|
    t.ignore_files << 'xl/calcChain.xml' << '[Content_Types].xml' << 'xl/_rels/workbook.xml.rels'

    wb.add_format(:bold, bold: true)
    ws = wb.add_worksheet 'Collapsed Rows'
    ws.set_column('A', 'A', width: 20)
    ws.set_selection 'A14:A14'
    (1..10).each do |i|
      ws.set_row i, level: ((i % 5).zero? ? 1 : 2), hide: true, collapse: (i % 5).zero?
    end
    ws.set_row 11, collapse: true

    write_common_data(ws)
  end

  test 'outline06' do |wb, t|
    t.ignore_files << 'xl/calcChain.xml' << '[Content_Types].xml' << 'xl/_rels/workbook.xml.rels'

    wb.add_format(:bold, bold: true)
    ws = wb.add_worksheet 'Outlined Rows'
    ws.outline_settings = { visible: false, symbols_below: false, symbols_right: false, auto_style: true }
    ws.set_column('A', 'A', width: 20)
    (1..10).each do |i|
      ws.set_row i, level: ((i % 5).zero? ? 1 : 2)
    end

    write_common_data(ws)
  end

  private

  def write_common_data(worksheet)
    worksheet
      .write_string('A1', 'Region', :bold)
      .write_string('A2', 'North')
      .write_string('A3', 'North')
      .write_string('A4', 'North')
      .write_string('A5', 'North')
      .write_string('A6', 'North Total', :bold)
      .write_string('B1', 'Sales', :bold)
      .write_number('B2', 1000)
      .write_number('B3', 1200)
      .write_number('B4', 900)
      .write_number('B5', 1200)
      .write_formula_num('B6', '=SUBTOTAL(9,B2:B5)', 4300, :bold)
      .write_string('A7', 'South')
      .write_string('A8', 'South')
      .write_string('A9', 'South')
      .write_string('A10', 'South')
      .write_string('A11', 'South Total', :bold)
      .write_number('B7', 400)
      .write_number('B8', 600)
      .write_number('B9', 500)
      .write_number('B10', 600)
      .write_formula_num('B11', '=SUBTOTAL(9,B7:B10)', 2100, :bold)
      .write_string('A12', 'Grand Total', :bold)
      .write_formula_num('B12', '=SUBTOTAL(9,B2:B10)', 6400, :bold)
  end
end
