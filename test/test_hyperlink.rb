# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestHyperlink < XlsxWriterTestCase
  test 'hyperlink03' do |wb|
    wb.unset_default_url_format
    ws1 = wb.add_worksheet
    ws2 = wb.add_worksheet

    ws1
      .write_url(0,  'A', 'http://www.perl.org/')
      .write_url(3,  'D', 'http://www.perl.org/')
      .write_url(7,  'A', 'http://www.perl.org/')
      .write_url(5,  'B', 'http://www.cpan.org/')
      .write_url(11, 'F', 'http://www.cpan.org/')

    ws2
      .write_url(1,  'C', 'http://www.google.com/')
      .write_url(4,  'C', 'http://www.cpan.org/')
      .write_url(6,  'C', 'http://www.perl.org/')
  end

  test 'hyperlink05' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet
      .write_url(0, 'A', 'http://www.perl.org/')
      .write_url(2, 'A', 'http://www.perl.org/', nil, string: 'Perl home')
      .write_url(4, 'A', 'http://www.perl.org/', string: 'Perl home', tooltip: 'Tool Tip')
      .write_url(6, 'A', 'http://www.cpan.org/', string: 'CPAN', tooltip: 'Download')
  end

  test 'hyperlink09' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet
      .write_url('A1', 'external:..\\foo.xlsx')
      .write_url('A3', 'external:..\\foo.xlsx#Sheet1!A1')
      .write_url('A5', 'external:\\\\VBOXSVR\\share\\foo.xlsx#Sheet1!B2',
                 string: 'J:\\foo.xlsx#Sheet1!B2')
  end

  test 'hyperlink14' do |wb, t|
    t.ignore_files = %w[xl/sharedStrings.xml]

    wb.add_format(:f, align: XlsxWriter::Format::ALIGN_CENTER)
    ws = wb.add_worksheet
    ws
      .write_string(0, 0, 'Perl Home', nil)
      .merge_range(3, 2, 4, 4, 'http://www.perl.org/', :f)
      .write_url(3, 2, 'http://www.perl.org/', format: :f, string: 'Perl Home')
  end

  test 'hyperlink18' do |wb|
    wb.max_url_length = 255
    ws = wb.add_worksheet
    wb.unset_default_url_format
    ws.write_url(0, 0, 'http://google.com/00000000001111111111222222222233333333334444444444555555555566666666666777777777778888888888999999999990000000000111111111122222222223333333333444444444455555555556666666666677777777777888888888899999999999000000000011111111112222222222x')
  end

  test 'hyperlink19' do |wb|
    pend
    wb.add_worksheet
      .write_url('A1', 'http://www.perl.com/')
      .write_formula_num('A1', '=1+1', 2)
    wb.sst.string_count = 0
  end

  test 'hyperlink20' do |wb, t|
    t.ignore_files = %w[xl/styles.xml]

    wb
      .add_format(:f1, underline: XlsxWriter::Format::UNDERLINE_SINGLE,
                       font_color: XlsxWriter::Format::COLOR_BLUE)
      .add_format(:f2, underline: XlsxWriter::Format::UNDERLINE_SINGLE,
                       font_color: XlsxWriter::Format::COLOR_RED)
      .add_worksheet
      .write_url(0, 0, 'http://www.python.org/1', :f1)
      .write_url(1, 0, 'http://www.python.org/2', :f2)
  end
end
