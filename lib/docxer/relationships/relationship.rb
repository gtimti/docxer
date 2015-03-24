module Docxer
  class Relationships
    class Relationship

      attr_accessor :id, :type, :target

      def initialize(id, type, target)
        @id = id
        @type = type
        @target = target
      end

      def build(xml)
        xml.Relationship('Id' => @id, 'Type' => @type, 'Target' => @target)
      end

    end
  end
end