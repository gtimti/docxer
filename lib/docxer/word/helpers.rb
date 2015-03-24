module Docxer
  module Word
    module Helpers

      def text(text, options={})
        p(options) do |p|
          p.text text
        end
      end

      def br(options={})
        element = Docxer::Word::Contents::Break.new(options)
        @content << element
        element
      end

      def p(options={}, &block)
        element = Docxer::Word::Contents::Paragraph.new(options, &block)
        @content << element
        element
      end

      def table_of_content(options={}, &block)
        toc = Docxer::Word::Contents::TableOfContent.new(options, &block)
        @content << toc
        toc
      end

      def table(options={}, &block)
        table = Docxer::Word::Contents::Table.new(options, &block)
        @content << table
        table
      end

    end
  end
end