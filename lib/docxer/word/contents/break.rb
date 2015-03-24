module Docxer
  module Word
    module Contents

      class Break

        attr_accessor :options
        def initialize(options={})
          @options = options
          @options[:times] ||= 1
        end

        def render(xml)
          if @options[:page]
            xml['w'].p do
              xml['w'].r do
                xml['w'].br( 'w:type' => "page" )
              end
            end
          else
            @options[:times].times do
              xml['w'].r do
                xml['w'].br
              end
            end
          end
        end

      end

    end
  end
end
