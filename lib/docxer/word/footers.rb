require 'docxer/word/footers/footer'

module Docxer
  module Word
    class Footers

      attr_accessor :footers, :counter
      def initialize
        @footers = []
        @counter = 0
      end

      def add(footer)
        @footers << footer
        footer
      end

      def sequence
        @counter += 1
      end

      def render(zip)
        @footers.each do |footer|
          footer.render(zip)
        end
      end

    end
  end
end