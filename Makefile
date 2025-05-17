install:
	uv sync

build:
	uv build

package-install:
	uv tool install dist/*.whl

num-convert:
	uv run docx-num-converter