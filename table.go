package excel

import (
	"github.com/xuri/excelize/v2"
)

type Data interface {
	Data() [][]any
}

type Position struct {
	Row    int
	Column int
}

type Header struct {
	Header string
	Size   float64
}

type TableData struct {
	Headers        []Header
	Rows           [][]any
	HeadersSize    float64
	RowsSize       float64
	Position       Position
	TablesOptions  []excelize.Table
	SlicersOptions []excelize.SlicerOptions
}

func NewTableData(headers []Header, rows Data, headersSize float64, rowsSize float64, position Position,
	tablesOptions []excelize.Table,
	slicerOptions []excelize.SlicerOptions,
) *TableData {
	return &TableData{
		Headers:        headers,
		HeadersSize:    headersSize,
		Rows:           rows.Data(),
		RowsSize:       rowsSize,
		TablesOptions:  tablesOptions,
		Position:       position,
		SlicersOptions: slicerOptions}
}

type Table struct {
	sheetName string
	table     TableData
}

func NewTable(sheetName string,
	table TableData,
) (*Table, error) {
	return &Table{sheetName: sheetName,
		table: table,
	}, nil
}

func (e *Table) SheetName() string {
	return e.sheetName
}

func (e *Table) Table() TableData {
	return e.table
}
