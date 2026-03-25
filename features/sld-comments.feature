Feature: Slide comments
  In order to annotate slides with review feedback
  As a developer using python-pptx
  I need to read and write comments on slides


  Scenario: Slide with no comments
    Given a slide with no comments
     Then slide.comments is an empty list


  Scenario: Read comments from a slide
    Given a slide with comments
     Then slide.comments has 2 items
      And the first comment text is "Comment one"
      And the first comment author is "Jane Smith"
      And the first comment has a timestamp
      And the second comment text is "Comment two"
      And the second comment author is "John Doe"


  Scenario: Add a comment to a blank slide
    Given a blank slide
     When I add a comment "Review this slide" by "Jane Smith"
     Then slide.comments has 1 items
      And the first comment text is "Review this slide"
      And the first comment author is "Jane Smith"
      And the first comment has a timestamp


  Scenario: Comments survive round-trip save and load
    Given a blank slide
     When I add a comment "Saved comment" by "Author A"
      And I save and reload the presentation and get the first slide
     Then slide.comments has 1 items
      And the first comment text is "Saved comment"
      And the first comment author is "Author A"


  Scenario: Add multiple comments by the same author
    Given a blank slide
     When I add a comment "First note" by "Same Author"
      And I add a comment "Second note" by "Same Author"
     Then slide.comments has 2 items
      And the first comment author is "Same Author"
      And the second comment author is "Same Author"
