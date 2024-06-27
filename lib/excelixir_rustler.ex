defmodule ExcelixirRustler do
  use Rustler, otp_app: :excelixir, crate: "excelixir_rustler"

  def read(_path), do: :erlang.nif_error(:nif_not_loaded)
  def get_sheet(_spreadsheet, _index), do: :erlang.nif_error(:nif_not_loaded)
  def get_sheet_by_name(_spreadsheet, _name), do: :erlang.nif_error(:nif_not_loaded)
  def get_cell(_sheet, _cell), do: :erlang.nif_error(:nif_not_loaded)

  def get_cell_value(_cell), do: :erlang.nif_error(:nif_not_loaded)
  def get_cell_value(_sheet, _cell_ref), do: :erlang.nif_error(:nif_not_loaded)

  def set_cell_value(_cell, _string), do: :erlang.nif_error(:nif_not_loaded)
  def set_cell_value(_sheet, _cell_ref, _value), do: :erlang.nif_error(:nif_not_loaded)

  def save(_spreadsheet, _path), do: :erlang.nif_error(:nif_not_loaded)

  ## Wrapper function to handle errors
  #def safe_get_cell_value(sheet, cell_ref) do
  #  try do
  #    get_cell_value(sheet, cell_ref)
  #  rescue
  #    ArgumentError ->
  #      {:error, "invalid_reference"}
  #  end
  #end
end
