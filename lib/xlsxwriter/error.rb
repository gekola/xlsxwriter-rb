# frozen_string_literal: true

module XlsxWriter
  ##
  # An error returned from libxlsxwriter. Numeric code is accessible via #code
  class Error < StandardError
    # Error code
    attr_reader :code

    def initialize(msg, code = nil)
      super(msg)
      @code = code
    end

    def inspect
      "#<#{self.class}: #{self}#{" (code: #{code})" if code}>"
    end
  end
end
