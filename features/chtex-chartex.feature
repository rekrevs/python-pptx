Feature: ChartEx chart creation
  In order to create modern Office 2016+ chart types in a presentation
  As a developer using python-pptx
  I need to add ChartEx charts to slides


  Scenario Outline: SlideShapes.add_chartex()
    Given a blank slide
     When I add a <chart-type> ChartEx chart with sample data
     Then the slide has a graphic frame containing a ChartEx chart
      And the ChartEx chart layout is <layout-id>

    Examples: ChartEx chart types
      | chart-type    | layout-id  |
      | Waterfall     | waterfall  |
      | Funnel        | funnel     |
      | Box & Whisker | boxWhisker |
      | Region Map    | regionMap  |
