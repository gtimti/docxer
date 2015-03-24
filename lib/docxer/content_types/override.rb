module Docxer
  class ContentTypes
    class Override
      
      attr_accessor :part_name, :content_type

      def initialize(part_name, content_type)
        @part_name = part_name
        @content_type = content_type
      end

      def build(xml)
        xml.Override('PartName' => @part_name, 'ContentType' => @content_type)
      end

    end
  end
end