install:
	uv sync --no-dev

install-dev:
	uv sync

build:
	uv build

package-install:
	uv tool install dist/*.whl

num-convert:
	uv run docx-num-converter