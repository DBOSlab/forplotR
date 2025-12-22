test_that(".arg_check_dir works correctly", {

  # Test with valid directory path (with trailing slash)
  expect_equal(.arg_check_dir("test/path/"), "test/path")

  # Test with valid directory path (without trailing slash)
  expect_equal(.arg_check_dir("test/path"), "test/path")

  # Test error with non-character input
  expect_error(
    .arg_check_dir(123),
    "The argument dir should be a character, not 'numeric'."
  )

  expect_error(
    .arg_check_dir(data.frame()),
    "The argument dir should be a character, not 'data.frame'."
  )

  # Test with empty string
  expect_equal(.arg_check_dir(""), "")

  # Test with just a slash (FIXED - should remove it)
  expect_equal(.arg_check_dir("/"), "")

  # Test with Windows-style backslashes (if applicable)
  # Note: Your function only handles forward slashes
  expect_equal(.arg_check_dir("test\\path\\"), "test\\path\\")

  # Test with mixed slashes at the end
  expect_equal(.arg_check_dir("test/path\\/"), "test/path\\")

  # Test with space and slash
  expect_equal(.arg_check_dir("test/path /"), "test/path ")
})

# Also need to test .arg_check_dir for different inputs
test_that(".arg_check_dir handles various inputs", {

  # Test with NA
  expect_error(
    .arg_check_dir(NA),
    "The argument dir should be a character, not 'logical'."
  )

  # Test with NULL
  expect_error(
    .arg_check_dir(NULL),
    "The argument dir should be a character, not 'NULL'."
  )
})

test_that(".arg_check_majorarea works for Brazil", {

  # Test with full state names (with diacritics)
  expect_equal(.arg_check_majorarea("São Paulo", "Brasil"), "São Paulo")
  expect_equal(.arg_check_majorarea("Paraná", "Brasil"), "Paraná")
  expect_equal(.arg_check_majorarea("Mato Grosso", "Brasil"), "Mato Grosso")

  # Test with state acronyms
  expect_equal(.arg_check_majorarea("SP", "Brasil"), "São Paulo")
  expect_equal(.arg_check_majorarea("PR", "Brasil"), "Paraná")
  expect_equal(.arg_check_majorarea("MT", "Brasil"), "Mato Grosso")

  # Test with state names without diacritics
  expect_equal(.arg_check_majorarea("Sao Paulo", "Brasil"), "São Paulo")
  expect_equal(.arg_check_majorarea("Parana", "Brasil"), "Paraná")
  expect_equal(.arg_check_majorarea("Espirito Santo", "Brasil"), "Espírito Santo")

  # Test error with invalid state
  expect_error(
    .arg_check_majorarea("InvalidState", "Brasil"),
    "'InvalidState' is not a Brazilian state."
  )

  # Test that non-Brazil countries return the input unchanged
  expect_equal(.arg_check_majorarea("California", "USA"), "California")
  expect_equal(.arg_check_majorarea("SomeRegion", "Colombia"), "SomeRegion")

  # Test missing argument
  expect_error(
    .arg_check_majorarea(country = "Brasil"),
    "'majorarea' is required."
  )
})

test_that(".arg_check_minorarea works correctly", {

  # Test that function doesn't return anything (just checks presence)
  expect_silent(.arg_check_minorarea("some value"))
  expect_silent(.arg_check_minorarea(123))
  expect_silent(.arg_check_minorarea(NULL))

  # Test missing argument
  expect_error(
    .arg_check_minorarea(),
    "'minorarea' is required."
  )
})

test_that(".arg_check_lat works correctly", {

  # Test with valid numeric input
  expect_equal(.arg_check_lat(-23.55), -23.55)
  expect_equal(.arg_check_lat(0), 0)
  expect_equal(.arg_check_lat(45.123), 45.123)

  # Test with character that can be converted to numeric
  expect_equal(.arg_check_lat("-23.55"), -23.55)
  expect_equal(.arg_check_lat("0"), 0)
  expect_equal(.arg_check_lat("45.123"), 45.123)

  # Test error with non-numeric character
  expect_error(
    .arg_check_lat("abc"),
    "'abc' must be a valid numeric latitude."
  )

  expect_error(
    .arg_check_lat("23.5N"),
    "'23.5N' must be a valid numeric latitude."
  )

  # Test missing argument
  expect_error(
    .arg_check_lat(),
    "'lat' is required."
  )
})

test_that(".arg_check_long works correctly", {

  # Test with valid numeric input
  expect_equal(.arg_check_long(-46.63), -46.63)
  expect_equal(.arg_check_long(0), 0)
  expect_equal(.arg_check_long(180), 180)

  # Test with character that can be converted to numeric
  expect_equal(.arg_check_long("-46.63"), -46.63)
  expect_equal(.arg_check_long("0"), 0)
  expect_equal(.arg_check_long("180"), 180)

  # Test error with non-numeric character
  expect_error(
    .arg_check_long("abc"),
    "'abc' must be a valid numeric longitude."
  )

  expect_error(
    .arg_check_long("46.63W"),
    "'46.63W' must be a valid numeric longitude."
  )

  # Test missing argument
  expect_error(
    .arg_check_long(),
    "'long' is required."
  )
})

test_that(".arg_check_alt works correctly", {

  # Test with valid numeric input
  expect_equal(.arg_check_alt(100), 100)
  expect_equal(.arg_check_alt(0), 0)
  expect_equal(.arg_check_alt(-50), -50)

  # Test with character that can be converted to numeric
  expect_equal(.arg_check_alt("100"), 100)
  expect_equal(.arg_check_alt("0"), 0)
  expect_equal(.arg_check_alt("-50"), -50)

  # Test with NULL (should return NULL)
  expect_equal(.arg_check_alt(NULL), NULL)

  # Test error with non-numeric character
  expect_error(
    .arg_check_alt("abc"),
    "'abc' must be a number."
  )

  expect_error(
    .arg_check_alt("100m"),
    "'100m' must be a number."
  )

  # Test with missing argument (NA)
  expect_true(is.na(.arg_check_alt(NA)))
})

test_that(".translate_country works correctly", {

  # Test with standard names (should return unchanged)
  expect_equal(.translate_country("Brazil"), "Brazil")
  expect_equal(.translate_country("Argentina"), "Argentina")
  expect_equal(.translate_country("Colombia"), "Colombia")

  # Test with Portuguese names
  expect_equal(.translate_country("Brasil"), "Brazil")
  expect_equal(.translate_country("Colômbia"), "Colombia")
  expect_equal(.translate_country("Paraguai"), "Paraguay")
  expect_equal(.translate_country("Uruguai"), "Uruguay")

  # Test with Spanish names
  expect_equal(.translate_country("España"), "Spain")
  expect_equal(.translate_country("México"), "Mexico")
  expect_equal(.translate_country("Bolívia"), "Bolivia")

  # Test with French names
  expect_equal(.translate_country("Brésil"), "Brazil")
  expect_equal(.translate_country("Colombie"), "Colombia")
  expect_equal(.translate_country("Pérou"), "Peru")

  # Test without diacritics
  expect_equal(.translate_country("Brazil"), "Brazil")  # Already without
  expect_equal(.translate_country("Colombia"), "Colombia")  # Already without
  expect_equal(.translate_country("Mexico"), "Mexico")  # Already without

  # Test error with unknown country
  expect_error(
    .translate_country("Neverland"),
    "Unknown country: Neverland"
  )

  # Test case sensitivity
  # expect_equal(.translate_country("BRAZIL"), "Brazil")
  # expect_equal(.translate_country("brazil"), "Brazil")
  # expect_equal(.translate_country("BrAzIl"), "Brazil")
})

test_that(".validate_plot_size works correctly", {

  # Test with valid plot sizes
  expect_silent(.validate_plot_size(0.2))
  expect_silent(.validate_plot_size(0.5))
  expect_silent(.validate_plot_size(1))

  # Test with invalid plot sizes
  expect_error(
    .validate_plot_size(0.1),
    "`plot_size` must be one of: 0.2, 0.5, or 1 hectare."
  )

  expect_error(
    .validate_plot_size(2),
    "`plot_size` must be one of: 0.2, 0.5, or 1 hectare."
  )

  # Test with numeric that's not exactly equal (floating point)
  expect_silent(.validate_plot_size(1.0))
  expect_error(.validate_plot_size(1.0001))
})

test_that(".validate_subplot_size works correctly", {

  # Test with valid subplot sizes
  expect_silent(.validate_subplot_size(10))
  expect_silent(.validate_subplot_size(20))
  expect_silent(.validate_subplot_size(25))

  # Test with invalid subplot sizes
  expect_error(
    .validate_subplot_size(5),
    "`subplot_size` must be one of: 10, 20, or 25 meters."
  )

  expect_error(
    .validate_subplot_size(30),
    "`subplot_size` must be one of: 10, 20, or 25 meters."
  )

  # Test with numeric that's not exactly equal
  expect_silent(.validate_subplot_size(10.0))
  expect_error(.validate_subplot_size(10.001))
})
