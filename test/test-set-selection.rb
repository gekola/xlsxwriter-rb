require_relative './xlsx-func-testcase'

class TestSetSelection < XlsxWriterTestCase
  test 'set_selection02' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    wb.add_worksheet { |ws| ws.set_selection(3,   2, 3,   2) }
    wb.add_worksheet { |ws| ws.set_selection(3,   2, 6,   6) }
    wb.add_worksheet { |ws| ws.set_selection(6,   6, 3,   2) }
    wb.add_worksheet { |ws| ws.set_selection(3, 'C', 3, 'C') }
    wb.add_worksheet { |ws| ws.set_selection(3, 'C', 6, 'G') }
    wb.add_worksheet { |ws| ws.set_selection(6, 'G', 3, 'C') }
  end
end
