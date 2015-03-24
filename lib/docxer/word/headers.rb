require 'docxer/word/headers/header'

module Docxer
  module Word
    class Headers

      attr_accessor :headers, :counter
      def initialize
        @headers = []
        @counter = 0
      end

      def add(header)
        @headers << header
        header
      end

      def sequence
        @counter += 1
      end

      def render(zip)
        @headers.each do |header|
          header.render(zip)
        end
      end

    end
  end
end