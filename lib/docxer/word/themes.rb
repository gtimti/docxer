require 'docxer/word/themes/theme'

module Docxer
  module Word
    class Themes

      attr_accessor :themes, :counter
      def initialize
        @themes = []
        @counter = 0
        
        add(Theme.new)
      end

      def add(theme)
        @counter += 1
        @themes << theme.set_sequence(@counter)
        theme
      end

      def render(zip)
        @themes.each do |theme|
          zip.put_next_entry("word/theme/theme#{theme.sequence}.xml")
          zip.write(theme.render)
        end
      end

    end
  end
end