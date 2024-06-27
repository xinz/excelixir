# Excelixir

**TODO: Add description**

## Installation

If [available in Hex](https://hex.pm/docs/publish), the package can be installed
by adding `excelixir` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:excelixir, "~> 0.1.0"}
  ]
end
```

## Usage

Here are some sample `NIF` functions API(see `ExcelixirRustler` module), need to update be with more friendly-wrapper API later.

```elixir
iex(1)> ref = ExcelixirRustler.read("/path/to/target/file.xlsx")
#Reference<0.679163101.1551237125.87864>
iex(2)> sheet = ExcelixirRustler.get_sheet(ref, 0)
#Reference<0.679163101.1551237125.87871>
iex(3)> ExcelixirRustler.set_cell_value(sheet, "A4", "testabc1")
:ok
iex(4)> ExcelixirRustler.save(ref, "/path/to/save/target-file.xlsx")
:ok
```

## Development

Local development be with the following rust version:

```
> rustc --version
rustc 1.79.0 (129f3b996 2024-06-10)
> rustup --version
rustup 1.27.1 (54dd3d00f 2024-04-24)
```
