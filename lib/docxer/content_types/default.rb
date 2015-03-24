module Docxer
  class ContentTypes
    class Default

      attr_accessor :extension, :content_type

      def initialize(extension, content_type)
        @extension = extension
        @content_type = content_type
      end

      def build(xml)
        xml.Default('Extension' => @extension, 'ContentType' => @content_type)
      end

    end
  end
end