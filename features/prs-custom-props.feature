Feature: Read and write custom document properties
  In order to store and retrieve arbitrary metadata for a presentation
  As a developer using python-pptx
  I need to access and assign custom document properties

  Scenario: Set and get a string custom property
     Given I have a new presentation
      When I set a custom property "Department" to string "Engineering"
       And I save the presentation
      Then the custom property "Department" of the saved presentation is "Engineering"

  Scenario: Set and get an int custom property
     Given I have a new presentation
      When I set a custom property "Revision" to int 42
       And I save the presentation
      Then the custom property "Revision" of the saved presentation is int 42

  Scenario: Set and get a bool custom property
     Given I have a new presentation
      When I set a custom property "Approved" to bool true
       And I save the presentation
      Then the custom property "Approved" of the saved presentation is bool true

  Scenario: Delete a custom property
     Given I have a new presentation
      When I set a custom property "Temp" to string "delete-me"
       And I delete the custom property "Temp"
       And I save the presentation
      Then the custom property "Temp" does not exist in the saved presentation
