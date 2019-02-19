require 'axlsx'
require 'xlsxwriter'
require 'benchmark'

Benchmark.bmbm do |x|
  x.report('xlsxwriter') do
    XlsxWriter::Workbook.new('/tmp/test.xlsx') do |wb|
      wb.add_worksheet do |ws|
        100_000.times { |i| ws.add_row([i]) }
      end
      wb.close
    end
  end

  x.report('axlsx') do
    Axlsx::Package.new do |p|
      p.workbook.add_worksheet do |ws|
        100_000.times { |i| ws.add_row([i]) }
      end
      p.serialize('/tmp/test.xlsx')
    end
  end
end
