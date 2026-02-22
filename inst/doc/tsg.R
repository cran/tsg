## ----include = FALSE----------------------------------------------------------
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)

## ----setup--------------------------------------------------------------------
library(tsg)

## -----------------------------------------------------------------------------
dim(person_record)
head(person_record)

## -----------------------------------------------------------------------------
person_record |> 
  generate_frequency(sex)

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(sex, age, marital_status)

## -----------------------------------------------------------------------------
person_record |>
  dplyr::group_by(sex) |>
  generate_frequency(marital_status)

## -----------------------------------------------------------------------------
person_record |>
  dplyr::group_by(sex) |>
  generate_frequency(marital_status, group_as_list = TRUE)

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(age, sort_value = TRUE)

person_record |>
  generate_frequency(age, sort_value = FALSE)

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    sex, 
    age, 
    marital_status, 
    # vector of variable names (character) to exclude from sorting
    sort_except = "age" 
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    marital_status,
    top_n = 3
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    marital_status, 
    top_n = 3,
    top_n_only = TRUE
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    employed,
    include_na = TRUE # default
  )

# Exclude NA values
person_record |>
  generate_frequency(
    employed,
    include_na = FALSE
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    seeing,
    hearing,
    walking,
    remembering,
    self_caring,
    communicating, 
    collapse_list = TRUE
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    seeing,
    hearing,
    walking,
    remembering,
    self_caring,
    communicating
  ) |> 
  collapse_list()

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    sex, 
    add_cumulative = TRUE, 
    add_cumulative_percent = TRUE 
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    marital_status,
    as_proportion = TRUE
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_frequency(
    marital_status,
    position_total = "top"
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_crosstab(marital_status, sex)

## -----------------------------------------------------------------------------
person_record |>
  generate_crosstab(
    sex,
    seeing,
    hearing,
    walking,
    remembering,
    self_caring,
    communicating
  )

## -----------------------------------------------------------------------------
person_record |>
  dplyr::group_by(sex) |>
  generate_crosstab(marital_status, employed)

## -----------------------------------------------------------------------------
person_record |>
  dplyr::group_by(sex) |>
  generate_crosstab(marital_status, employed, group_as_list = TRUE)

## -----------------------------------------------------------------------------
person_record |>
  generate_crosstab(
    marital_status,
    sex,
    percent_by_column = TRUE
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_crosstab(
    marital_status,
    sex,
    as_proportion = TRUE
  )

## -----------------------------------------------------------------------------
person_record |>
  generate_crosstab(
    marital_status,
    sex,
    position_total = "top"
  )

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_frequency(sex) |>
#   write_xlsx(path = "table-01.xlsx")

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_crosstab(marital_status, sex) |>
#   add_table_title("Marital Status by Sex") |>
#   add_table_subtitle("Sample dataset: person_record") |>
#   write_xlsx(path = "table-02.xlsx")

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_crosstab(marital_status, sex) |>
#   add_table_title("Marital Status by Sex") |>
#   add_table_subtitle("Sample dataset: person_record") |>
#   add_source_note("Source: person_record dataset") |>
#   add_footnote("This is a footnote for the table") |>
#   write_xlsx(path = "table-03.xlsx")

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_crosstab(marital_status, sex) |>
#   write_xlsx(
#     path = "table-03.xlsx",
#     table_title = "Marital Status by Sex",
#     table_subtitle = "Sample dataset: person_record",
#     source_note = "Source: person_record dataset",
#     footnotes = "This is a footnote for the table"
#   )

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_frequency(sex) |>
#   add_facade(
#     table.offsetRow = 2,
#     table.offsetCol = 1
#   ) |>
#   write_xlsx(
#     path = "table-04.xlsx",
#     # Using built-in facade
#     facade = get_tsg_facade("yolo")
#   )

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_frequency(sex) |>
#   write_xlsx(
#     path = "table-05.xlsx",
#     # Using built-in facade
#     facade = get_tsg_facade("yolo")
#   )

## ----eval=FALSE---------------------------------------------------------------
# person_record |>
#   generate_frequency(sex) |>
#   generate_output(path = "table-06.xlsx")

