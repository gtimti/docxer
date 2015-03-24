module Docxer
  module Word
    class Medias
      class Media

        attr_accessor :id, :sequence, :file
        def initialize(id, file, options)
          @id = id
          @file = file
          self
        end

        def set_sequence(sequence)
          @sequence = sequence
          self
        end

        def type
          "#_x0000_t75"
        end

        def target
          "media/image#{id}.png"
        end

        def content_type
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        end

        def uniq_id
          "_x0000_i102#{id}"
        end

      end
    end
  end
end