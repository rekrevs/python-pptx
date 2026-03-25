Feature: Shape effects
  In order to adjust visual effects on shapes
  As a developer using python-pptx
  I need properties and methods on effect format objects


  Scenario Outline: ShadowFormat.inherit getter
    Given a ShadowFormat object that <inherits-or-not> as shadow
     Then shadow.inherit is <value>

    Examples: shadow inheritance cases
      | inherits-or-not  | value |
      | inherits         | True  |
      | does not inherit | False |


  Scenario Outline: ShadowFormat.inherit setter
    Given a ShadowFormat object that <inherits-or-not> as shadow
     When I assign <new-value> to shadow.inherit
     Then shadow.inherit is <value>

    Examples: shadow inheritance assignment cases
      | inherits-or-not  | new-value | value |
      | inherits         | False     | False |
      | does not inherit | True      | True  |
      | inherits         | None      | False |
      | inherits         | True      | True  |
      | does not inherit | None      | False |
      | does not inherit | False     | False |


  Scenario: ShadowFormat shadow_type getter
    Given a shape with an outer shadow as shadow_shape
     Then shadow.shadow_type is "outer"


  Scenario: ShadowFormat angle getter
    Given a shape with an outer shadow as shadow_shape
     Then shadow.angle is 90.0


  Scenario: ShadowFormat blur_radius getter
    Given a shape with an outer shadow as shadow_shape
     Then shadow.blur_radius is 40000


  Scenario: ShadowFormat distance getter
    Given a shape with an outer shadow as shadow_shape
     Then shadow.distance is 23000


  Scenario: ShadowFormat color getter
    Given a shape with an outer shadow as shadow_shape
     Then shadow.color.rgb is "000000"
