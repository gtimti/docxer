module Docxer
  module Word
    module Contents

      class Image

        attr_accessor :media, :options
        def initialize(media, options={})
          @media = media
          @media.file.rewind
          @options = options
        end

        def styles
          if @options[:style]
            @options[:style].collect{|k, v| [k, v].join(':')}.join(';')
          end
        end

        def render(xml)
          xml['w'].r do
            xml['w'].rPr do
              xml['w'].noProof
            end
            xml['w'].pict do
              xml['v'].shape( 'id' => @media.uniq_id, 'type' => @media.type, 'style' => styles ) do
                xml['v'].imagedata( 'r:id' => @media.sequence, 'o:title' => @options[:title] )
              end
            end
            xml['w'].rPr do
              xml['w'].noProof
            end
          end
        end
      end
    end
  end
end
