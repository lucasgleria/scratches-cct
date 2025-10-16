import os
from playwright.sync_api import sync_playwright, expect

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # This script is injected before the page's scripts run.
        # It creates a mock of the `google.script.run` API.
        mock_script = """
        window.google = {
            script: {
                run: {
                    _success: null, _failure: null, _finally: null,
                    withSuccessHandler: function(handler) { this._success = handler; return this; },
                    withFailureHandler: function(handler) { this._failure = handler; return this; },
                    withFinally: function(handler) { this._finally = handler; return this; },
                    getSpreadsheets: function() {
                        const result = {
                            ok: true,
                            payload: [
                                { id: '1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0', name: 'CCT Teste' },
                                { id: '1qgP2RIXiA5cjO-EdSUjti11r6jBXVJR0PeMotHoBAA4', name: 'CCT Teste 2' }
                            ]
                        };
                        if (this._success) { this._success(result); }
                        if (this._finally) { this._finally(); }
                    },
                    checkHouseExists: function(spreadsheetId, mawb, house) {
                        const result = { ok: true, payload: { exists: true } };
                        if (this._success) { this._success(result); }
                        if (this._finally) { this._finally(); }
                    },
                    reset: function() {
                        this._success = null;
                        this._failure = null;
                        this._finally = null;
                    }
                }
            }
        };
        """
        page.add_init_script(mock_script)

        html_file_path = os.path.abspath("index.html")
        page.goto(f"file://{html_file_path}")

        # Manually trigger DOMContentLoaded to ensure the app's scripts run
        page.evaluate("document.dispatchEvent(new Event('DOMContentLoaded', { bubbles: true, cancelable: true }));")

        # Wait for the mock to populate the spreadsheet dropdown
        page.wait_for_selector("#spreadsheet option[value='1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0']", state='attached')

        # Fill out the form
        page.select_option("#spreadsheet", "1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0")
        page.fill("#mawb", "12345678901")
        page.fill("#newHouse", "DUPLICATE-HOUSE")

        # Click the "Adicionar" button to trigger the validation
        page.click("#addHouseBtn")

        # Add a small delay to allow the async call to complete
        page.wait_for_timeout(500)

        # The core of the test: assert that the error message is visible
        error_locator = page.locator("#houseError")
        expect(error_locator).to_be_visible()
        expect(error_locator).to_have_text("Este HOUSE já existe na planilha para o MAWB informado.")

        # Capture the result in a screenshot
        page.screenshot(path="jules-scratch/verification/verification.png")

        browser.close()

if __name__ == "__main__":
    run()