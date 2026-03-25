Feature: Theme access
  In order to control the visual identity of a presentation
  As a developer using python-pptx
  I need to read and write theme colors and fonts


  Scenario: Read accent1 color from theme
    Given a presentation with a default theme
     Then the theme color scheme accent1 is "4F81BD"


  Scenario: Read all 12 theme colors
    Given a presentation with a default theme
     Then the theme color scheme dark1 is "000000"
      And the theme color scheme light1 is "FFFFFF"
      And the theme color scheme dark2 is "1F497D"
      And the theme color scheme light2 is "EEECE1"
      And the theme color scheme accent1 is "4F81BD"
      And the theme color scheme accent2 is "C0504D"
      And the theme color scheme accent3 is "9BBB59"
      And the theme color scheme accent4 is "8064A2"
      And the theme color scheme accent5 is "4BACC6"
      And the theme color scheme accent6 is "F79646"
      And the theme color scheme hyperlink is "0000FF"
      And the theme color scheme followed_hyperlink is "800080"


  Scenario: Write accent1 color to theme
    Given a presentation with a default theme
     When I set the theme color scheme accent1 to "FF0000"
     Then the theme color scheme accent1 is "FF0000"


  Scenario: Theme color changes survive round-trip
    Given a presentation with a default theme
     When I set the theme color scheme accent1 to "AABB00"
      And I save and reload the presentation
     Then the theme color scheme accent1 is "AABB00"


  Scenario: Read major font from theme
    Given a presentation with a default theme
     Then the theme font scheme major font is "Calibri"


  Scenario: Read minor font from theme
    Given a presentation with a default theme
     Then the theme font scheme minor font is "Calibri"


  Scenario: Write major font to theme
    Given a presentation with a default theme
     When I set the theme font scheme major font to "Arial"
     Then the theme font scheme major font is "Arial"


  Scenario: Write minor font to theme
    Given a presentation with a default theme
     When I set the theme font scheme minor font to "Helvetica"
     Then the theme font scheme minor font is "Helvetica"


  Scenario: Theme font changes survive round-trip
    Given a presentation with a default theme
     When I set the theme font scheme major font to "Comic Sans MS"
      And I set the theme font scheme minor font to "Tahoma"
      And I save and reload the presentation
     Then the theme font scheme major font is "Comic Sans MS"
      And the theme font scheme minor font is "Tahoma"
