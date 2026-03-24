Feature: Bullet and numbering formatting
  In order to control bullet and numbering appearance on paragraphs
  As a developer using python-pptx
  I need read/write access to bullet properties on paragraphs


  Scenario: Get bullet type for paragraph with no explicit bullet
    Given a paragraph with no explicit bullet setting
     Then paragraph.bullet.type is None


  Scenario: Set bullet type to character bullet
    Given a paragraph with no explicit bullet setting
     When I assign "char" to paragraph.bullet.type
     Then paragraph.bullet.type is "char"
      And paragraph.bullet.char is not None


  Scenario: Set bullet type to auto-number
    Given a paragraph with no explicit bullet setting
     When I assign "auto_num" to paragraph.bullet.type
     Then paragraph.bullet.type is "auto_num"
      And paragraph.bullet.auto_num_type is "arabicPeriod"


  Scenario: Set bullet type to none (suppress bullets)
    Given a paragraph with no explicit bullet setting
     When I assign "none" to paragraph.bullet.type
     Then paragraph.bullet.type is "none"


  Scenario: Remove explicit bullet setting (inherit)
    Given a paragraph with a character bullet
     When I assign None to paragraph.bullet.type
     Then paragraph.bullet.type is None


  Scenario: Set bullet character
    Given a paragraph with no explicit bullet setting
     When I assign "-" to paragraph.bullet.char
     Then paragraph.bullet.char is "-"
      And paragraph.bullet.type is "char"


  Scenario: Set bullet auto-numbering type
    Given a paragraph with no explicit bullet setting
     When I assign "romanLcPeriod" to paragraph.bullet.auto_num_type
     Then paragraph.bullet.auto_num_type is "romanLcPeriod"
      And paragraph.bullet.type is "auto_num"


  Scenario: Set bullet font
    Given a paragraph with no explicit bullet setting
     When I assign "Wingdings" to paragraph.bullet.font
     Then paragraph.bullet.font is "Wingdings"


  Scenario: Remove bullet font
    Given a paragraph with bullet font set to "Arial"
     When I assign None to paragraph.bullet.font
     Then paragraph.bullet.font is None


  Scenario: Set bullet size
    Given a paragraph with no explicit bullet setting
     When I assign 14.0 to paragraph.bullet.size
     Then paragraph.bullet.size is 14.0


  Scenario: Remove bullet size
    Given a paragraph with bullet size set to 12.0
     When I assign None to paragraph.bullet.size
     Then paragraph.bullet.size is None


  Scenario: Set bullet color to RGB
    Given a paragraph with no explicit bullet setting
     When I assign RGBColor(0xFF, 0x00, 0x00) to paragraph.bullet.color.rgb
     Then paragraph.bullet.color.rgb is RGBColor(0xFF, 0x00, 0x00)
