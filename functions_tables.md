# Functions and what they do

Here is a plain-English guide to the functions in `muffin_recipe_builder_v5.py`. The function names are kept as they appear in the code so they can be looked up there.

| Section | Function | What it does |
|---|---|---|
| Ingredients file | `read_csv_text(path)` | Opens the ingredients file and reads its contents. |
| Ingredients file | `load_ingredients_from_csv(path)` | Reads the ingredients file and organizes the information for the app. If the file cannot be used, it reports a problem. |
| Ingredient details | `Ingredient.__init__(name, data)` | Sets up the name and information for one ingredient. |
| Ingredient details | `Ingredient.available_units()` | Lists the measuring options available for an ingredient. |
| Ingredient details | `Ingredient.grams_for(amount, unit)` | Converts an amount of an ingredient into grams when that measuring option is available. |
| Ingredient details | `Ingredient.macros_for_grams(grams)` | Works out the nutrition amounts for a given weight of an ingredient. |
| Starting the app | `MacroCalculatorApp.__init__()` | Starts the app, loads the ingredients, builds the screen, and shows a warning if there are problems with the ingredients file. |
| Starting the app | `_load_ingredients()` | Loads the ingredients and keeps track of any problems found in the ingredients file. |
| Starting the app | `_show_startup_problems()` | Shows a warning about problems found while loading the ingredients. |
| Screen and look | `_build_ui()` | Builds the app screen and connects its buttons, menus, and keyboard controls to features such as ingredient search, recipes, Excel files, and the light or dark look. |
| Screen and look | `_toggle_appearance()` | Switches the app between its light and dark looks. |
| Screen and look | `_style_treeview()` | Updates how the recipe table looks to match the current light or dark setting. |
| Finding ingredients | `_populate_results(names)` | Shows the ingredients found as a list of buttons. |
| Finding ingredients | `_filter_ingredients(event=None)` | Narrows the ingredient list to match what has been typed into the search box. |
| Finding ingredients | `_move_cursor(delta)` | Moves the highlight up or down through the ingredient results when the arrow keys are used. |
| Finding ingredients | `_highlight_cursor(index)` | Highlights an ingredient in the results and scrolls it into view. |
| Finding ingredients | `_confirm_cursor(event=None)` | Chooses the highlighted ingredient when Enter is pressed. |
| Finding ingredients | `_select_ingredient(name)` | Chooses an ingredient and shows the measuring options available for it. |
| Finding ingredients | `_sync_search_to_selection(name)` | Updates the search box to match the ingredient that was chosen. |
| Finding ingredients | `_reset_selection()` | Clears the current choice and search text, then shows the full ingredient list again. |
| Finding ingredients | `_highlight_selected_row()` | Highlights the chosen ingredient in its display. |
| Finding ingredients | `_select_all_in_search(event=None)` | Selects all the text in the search box when it is clicked into. |
| Building a recipe | `_add_ingredient()` | Checks the entered amount and measuring option, works out the ingredient's nutrition amounts, and adds it to the recipe. |
| Building a recipe | `_show_context_menu(event)` | Opens the recipe-row menu when the user right-clicks, selecting the row first if needed. |
| Building a recipe | `_remove_selected()` | Removes the chosen ingredient rows from the recipe. |
| Building a recipe | `_refresh_table()` | Updates the recipe table after ingredients are added, removed, or loaded. |
| Building a recipe | `_refresh_totals()` | Updates the recipe's nutrition totals and amounts per serving. |
| Choosing an Excel file | `_choose_existing_workbook()` | Lets the user choose an existing Excel file and remembers which file to use. |
| Choosing an Excel file | `_create_new_workbook()` | Lets the user choose where to save a new Excel file and remembers that location. |
| Choosing an Excel file | `_ask_new_or_existing()` | Asks whether the user wants to make a new Excel file or choose one they already have. |
| Choosing an Excel file | `_ask_new_or_existing.pick(value)` | Records the user's choice and closes the question window. |
| Choosing an Excel file | `_pick_workbook_button()` | When the “Load Workbook” button is clicked, asks whether to use a new or existing Excel file and follows that choice. |
| Loading a recipe | `_ask_pick_sheet(sheet_names)` | Asks which sheet in the Excel file the user wants to load. |
| Loading a recipe | `_ask_pick_sheet.confirm(name)` | Records the chosen sheet and closes the selection window. |
| Loading a recipe | `_parse_recipe_sheet(ws)` | Reads a recipe from an Excel sheet, including servings, and notes ingredient names it could not match. |
| Loading a recipe | `_load_recipe_from_excel()` | Loads a recipe from the chosen Excel file and updates the recipe display and totals. |
| Saving a recipe | `_save_recipe()` | Saves the recipe and its summary in an Excel file. |

`CsvLoadError` is not a function. It is the name used when the ingredients file cannot be loaded.
