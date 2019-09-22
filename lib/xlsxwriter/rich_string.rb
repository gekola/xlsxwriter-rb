# frozen_string_literal: true

module XlsxWriter
  ##
  # RichString is a multipart string consting of sub-strings with format
  class RichString
    attr_reader :workbook, :parts

    ##
    # call-seq:
    #   RichString.new(workbook, parts)
    #   RichString.new(workbook)
    #
    # Creates rich string that could be written to the +workbook+.
    # +parts+ is an array of pairs string + format. If no format is needed
    # for the part single string can be specified instead.
    # For example: +['This is ', ['bold', :bold], [' and this is'], ['italic', :italic]]+
    #
    # If parts is not specified it could be populated later with #add_part.
    #
    # Please note that on writing additional validations are applied (e.g. string should not be empty).
    # Please consult libxlsxwriter documentation for additional details.
    def initialize(workbook, parts = [])
      @workbook = workbook
      @parts = parts.map! do |(str, format)|
        make_part(str, format)
      end
    end

    ##
    # call-seq:
    #   rs.add_part(str)
    #   rs.add_part(str, format)
    #
    # Adds part (string +str+ formatted with +format+) to the RichString +rs+.
    #
    # Please note, that no parts can be added to a rich string that has already been written once.
    def add_part(str, format = nil)
      raise Error, 'Cannot modify as the RichString has already been written to a workbook.' if cached?

      parts << make_part(str, format)

      self
    end

    def <<(part)
      add_part(*Array(part))
    end

    def inspect
      "#<#{self.class}:#{to_s.inspect}, @cached=#{@cached}>"
    end

    def to_s
      parts.map(&:first)
    end

    private

    def make_part(str, format)
      [
        str.frozen? ? str.to_str : str.to_str.clone.freeze,
        (format.to_sym if format)
      ].freeze
    end
  end
end
