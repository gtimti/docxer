module Docxer
  module Word
    module Contents

      class Paragraph

        attr_accessor :content, :options
        def initialize(options={})
          @content = []
          @options = options

          if block_given?
            yield self
          else

          end
        end

        def render(xml)
          xml['w'].p do
            xml['w'].pPr do
              xml['w'].jc( 'w:val' => @options[:align]) if @options[:align]
              if options[:ul]
                xml['w'].numPr do
                  xml['w'].ilvl( 'w:val' => 0 )
                  xml['w'].numId( 'w:val' => 1 )
                end
              end
            end
            @content.each do |element|
              element.render(xml)
            end
          end
        end

        def text(text, options={})
          options = @options.merge(options)
          text = Docxer::Word::Contents::Text.new(text, options)
          @content << text
          text
        end

        def br(options={})
          br = Docxer::Word::Contents::Break.new(options)
          @content << br
          br
        end

        def image(image, options={})
          img = Docxer::Word::Contents::Image.new(image, options)
          @content << img
          img
        end

      end

    end
  end
end
