module Docxer
  module Word
    module Contents

      class TableOfContent

        attr_accessor :options, :rows
        def initialize(options={})
          @options = options
          @rows = []

          if block_given?
            yield self
          else

          end
        end

        def add_row(text, page=nil, options={})
          row = Row.new(text, page, options)
          @rows << row
          row
        end

        def render(xml)
          @rows.each do |row|
            row.render(xml)
          end
        end


        class Row

          attr_accessor :text, :page, :options
          def initialize(text, page=nil, options={})
            @text = text
            @page = page
            @options = options
          end

          def render(xml)
            if @page.nil?
              xml['w'].p do
                xml['w'].pPr do
                  xml['w'].spacing( 'w:before' => "0", 'w:after' => "160")
                end
                xml['w'].r do
                  xml['w'].tab
                  xml['w'].t @text
                end
              end
            else
              xml['w'].p do
                xml['w'].pPr do
                  xml['w'].pStyle( 'w:val' => "ListParagraph") # TODO: Apply styles
                  xml['w'].spacing( 'w:before' => "0", 'w:after' => "160")
                  xml['w'].tabs do
                    xml['w'].tab( 'w:val' => "right", 'w:leader' => @options[:leader] || 'dot', 'w:pos' => @options[:pos] || 9360 )
                  end
                end
                xml['w'].r do
                  xml['w'].t @text
                end
                xml['w'].r do
                  xml['w'].tab
                  xml['w'].t( @page, 'xml:space' => "preserve" )
                end
              end
            end
          end
        end
      end
    end
  end
end
