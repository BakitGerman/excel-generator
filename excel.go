package excel-generator

import (
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Excel[T any] interface {
	SheetName() string
	Table() TableData
}

/*
generate excel horizontal or vertical table
`column1` `row1`
`column2` `row2`
or
`column1` `column2`
`row1`    `row2`
*/
func ExcelGenerator[T any](data Excel[T], file *excelize.File, filename string, rotate bool) (*excelize.File, error) {

	// prepare data
	sheetName := data.SheetName()
	table := data.Table()
	colOffset := table.Position.Column
	rowOffset := table.Position.Row

	// set headers
	for i, header := range table.Headers {
		colName, err := excelize.ColumnNumberToName(colOffset + i)
		if err != nil {
			return nil, err
		}
		if err := file.SetCellStr(sheetName, colName+strconv.Itoa(rowOffset), header.Header); err != nil {
			return nil, err
		}
		if err := file.SetColWidth(sheetName, colName, colName, header.Size); err != nil {
			return nil, err
		}
	}
	if err := file.SetRowHeight(sheetName, rowOffset, table.HeadersSize); err != nil {
		return nil, err
	}
	// set rows
	for i := range table.Rows {
		if err := file.SetRowHeight(sheetName, rowOffset+i+1, table.RowsSize); err != nil {
			return nil, err
		}
		for j, columns := range table.Rows[i] {
			colName, err := excelize.ColumnNumberToName(colOffset + j)
			if err != nil {
				return nil, err
			}
			switch v := any(columns).(type) {
			case string:
				if len(v) > 0 && v[0] == '=' {
					if err := file.SetCellFormula(sheetName, colName+strconv.Itoa(rowOffset+i+1), v); err != nil {
						return nil, err
					}
				} else {
					if err := file.SetCellValue(sheetName, colName+strconv.Itoa(rowOffset+i+1), columns); err != nil {
						return nil, err
					}
				}
			default:
				if err := file.SetCellValue(sheetName, colName+strconv.Itoa(rowOffset+i+1), columns); err != nil {
					return nil, err
				}
			}
		}
	}

	// set tables
	for _, tableOpt := range table.TablesOptions {
		if err := file.AddTable(sheetName, &tableOpt); err != nil {
			return nil, err
		}
	}
	// set slicers
	for _, options := range table.SlicersOptions {
		if err := file.AddSlicer(sheetName, &options); err != nil {
			return nil, err
		}
	}

	return file, nil
}
