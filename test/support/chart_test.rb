# frozen_string_literal: true

module ChartTest
  def chart_test(fname, type, data = self::DATA, ref_file_name: nil, &_block)
    test fname, ref_file_name: ref_file_name do |wb|
      wb.add_worksheet do |ws|
        data.each { |row| ws.add_row row }

        wb.add_chart(type) do |chart|
          yield chart, ws

          ws.insert_chart 'E9', chart
        end
      end
    end
  end
end
