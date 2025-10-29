import asyncio
from playwright.sync_api import sync_playwright, Page, expect
import pathlib

def run_verification(page: Page):
    # Get the absolute path to the index.html file
    index_html_path = str(pathlib.Path('./index.html').resolve())

    # Go to the local HTML file
    page.goto(f'file://{index_html_path}')

    # 1. Enter MAWB
    mawb_input = page.get_by_label("MAWB")
    mawb_input.fill("12345678901")
    # Click somewhere else to trigger validation styling
    page.locator('h2').first.click()

    # 2. Add HOUSEs in bulk
    house_textarea = page.locator('#newHouses')
    house_textarea.fill("HOUSE1\nHOUSE2\nHOUSE3")
    page.get_by_role("button", name="Adicionar").click()

    # 3. Select a HOUSE
    house2 = page.get_by_text("HOUSE2")
    expect(house2).to_be_visible()
    house2.click()

    # Verify the main data section is now visible
    main_data_section = page.locator("#dataSection")
    expect(main_data_section).to_be_visible()

    # 4. Enter data for the selected HOUSE (and trigger auto-save with blur)
    entregas_input = page.get_by_placeholder("Digite uma entrega")
    entregas_input.fill("Entrega 1")
    entregas_input.press("Tab") # Simulate blur to trigger auto-save

    # Verify the value was saved and is now displayed
    expect(page.get_by_text("Entrega 1")).to_be_visible()

    # 5. Edit a HOUSE name
    house1 = page.get_by_text("HOUSE1")
    house1.dblclick()

    # After double-clicking, an input field should appear
    edit_input = page.locator('.house-list-item input[type="text"]')
    expect(edit_input).to_be_visible()
    edit_input.fill("HOUSE1_EDITED")
    edit_input.press("Enter")

    # Verify the name has been updated
    expect(page.get_by_text("HOUSE1_EDITED")).to_be_visible()

    # 6. Toggle the refrigerated cargo switch
    toggle_container = page.locator("div.toggle-container", has_text="Carga de Geladeira")
    refrigerated_cargo_switch = toggle_container.locator("label.switch")
    checkbox = page.locator("#fridgeToggle_HOUSE2")

    # Expect the checkbox to become checked after we click the switch
    # This combines the action and assertion, making it more robust to timing issues
    expect(checkbox).not_to_be_checked() # Ensure it's not checked initially
    refrigerated_cargo_switch.click()
    expect(checkbox).to_be_checked()

    # The options for fridge type should become visible
    fridge_options = page.locator('#fridgeOptions_HOUSE2')
    expect(fridge_options).to_be_visible()

    # Select an option
    fridge_type_select = page.locator('#fridgeType_HOUSE2')
    fridge_type_select.select_option("FRI")

    # The special data section for 'Observações' should now show the value
    observacoes_section = page.locator(".data-section", has_text="Observações")
    expect(observacoes_section.get_by_text("FRI")).to_be_visible()

    # 7. Take a screenshot of the final state
    page.screenshot(path="jules-scratch/verification/verification.png")

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        run_verification(page)
        browser.close()

if __name__ == "__main__":
    main()
