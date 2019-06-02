# frozen_string_literal: true

module ChartTest
  def chart_test(fname, type, data = self::DATA, &block)
    test fname do |wb|
      wb.add_worksheet do |ws|
        data.each { |row| ws.add_row row }

        wb.add_chart(type) do |chart|
          yield chart

          ws.insert_chart 'E9', chart
        end
      end
    end
  end
end
