# GitHub Copilot Instructions for Python Project

## Goal

Generate high-quality, efficient, maintainable, and robust Python code following modern best practices. Assume development is done using Visual Studio Code on Windows.

## General Guidelines

* **Language:** Python 3.10+
* **Style:** Adhere strictly to PEP 8. Use `black` for formatting and `ruff` or `flake8` for linting.
* **Clarity:** Prioritize readability and maintainability. Use clear variable/function names and add concise comments where logic isn't obvious.
* **Type Hinting:** Use type hints for all function signatures and complex variable assignments. Utilize the `typing` module.
* **Docstrings:** Generate Google-style docstrings for all modules, classes, functions, and methods. Include `Args:`, `Returns:`, and `Raises:`.

## Code Generation

* **Optimization:**
    * Prefer built-in functions and standard library modules where possible.
    * Use list comprehensions, generator expressions, and dictionary comprehensions for concise and efficient iterations.
    * Be mindful of algorithmic complexity (Big O notation). Suggest efficient algorithms for the task.
    * Avoid premature optimization; focus on clarity first unless performance is critical.
* **Accuracy:**
    * Pay close attention to the surrounding code context and comments.
    * If requirements are ambiguous, generate a simple, clear implementation and add a comment suggesting potential alternatives or asking for clarification.
    * Generate code that directly addresses the prompt or comment preceding it.

## Directory Structure

Assume and maintain the following project structure:


project-root/
├── .github/
│   └── copilot-instructions.md
├── src/
│   └── uco_to_udo_recon/
│       ├── init.py
│       ├── main.py
│       ├── core/
│       ├── modules/
│       └── utils/
├── tests/
│   ├── init.py
│   └── test_*.py
├── docs/
│   └── ...
├── data/          # Optional: For data files
├── scripts/       # Optional: For helper scripts
├── .gitignore
├── pyproject.toml # Or requirements.txt
└── README.md


* Place core application logic within `src/uco_to_udo_recon/`.
* Place unit and integration tests within `tests/`. Test files should mirror the structure of the `src/` directory.
* Use relative imports within the `src/` directory (e.g., `from .core import ...`).

## Error Handling

* **Specific Exceptions:** Catch specific exceptions rather than generic `Exception`.
* **Custom Exceptions:** Define custom exception classes for application-specific errors when appropriate.
* **`try...except...finally`:** Use `finally` blocks for cleanup operations (e.g., closing files or network connections).
* **Context Managers:** Use the `with` statement for resource management (files, locks, connections).
* **Error Messages:** Provide clear and informative error messages.

## Debugging

* **Variable Names:** Use descriptive variable names.
* **Intermediate Variables:** Don't shy away from using intermediate variables to clarify steps in complex calculations or logic.
* **Pure Functions:** Prefer pure functions (functions whose output depends only on their input and have no side effects) where possible, as they are easier to test and debug.
* **Assertions:** Use `assert` statements for sanity checks during development (but be aware they can be disabled).

## Logging

* **Standard Library:** Use the built-in `logging` module.
* **Configuration:** Configure logging early in the application's entry point (`main.py`). Consider configuration via file (`logging.conf`) or dictionary (`logging.config.dictConfig`).
* **Levels:** Use appropriate logging levels (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`).
* **Context:** Include relevant context in log messages (e.g., function names, relevant variable values).
* **Avoid `print()`:** Replace `print()` statements used for debugging or status updates with appropriate `logging` calls.

## Testing

* **Framework:** Use `pytest` as the testing framework.
* **Coverage:** Generate tests that aim for high code coverage.
* **Fixtures:** Utilize `pytest` fixtures for setting up and tearing down test states.
* **Mocking:** Use `unittest.mock` (or `pytest-mock`) for mocking dependencies.
* **Test Naming:** Test functions should start with `test_`.

## Security

* **Input Validation:** Always validate and sanitize external input (user input, API responses, file contents).
* **Secrets Management:** Do not hardcode secrets (API keys, passwords). Use environment variables or a dedicated secrets management tool. Suggest placeholders like `os.getenv("API_KEY")`.
* **Dependencies:** Be mindful of third-party library security vulnerabilities. (While Copilot can't check this directly, generating standard dependency files like `pyproject.toml` helps tooling).
