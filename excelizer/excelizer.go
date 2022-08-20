package excelizer

import (
	"container/list"
	"errors"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

var ErrEmptyBreakpoints = errors.New("empty breakpoints")

type Breakpoint struct {
	value    string
	cellName string
}

type Sheet struct {
	file        *excelize.File
	breakpoints *list.List
}

func OpenSheet(filename string) (*Sheet, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("OpenSheet: %v", err)
	}

	return &Sheet{file: f}, nil
}

func (s *Sheet) ShowBreakpoints() {
	bp := s.breakpoints.Front()

	for i := s.breakpoints.Len(); i > 0; i-- {

		fmt.Println(*bp.Value.(*Breakpoint))
		bp = bp.Next()
	}
}

func (s *Sheet) CloseSheet() error {
	return s.file.Close()
}

func (s *Sheet) SetBreakpoints(sheet string, bp ...string) error {
	if len(bp) == 0 {
		return ErrEmptyBreakpoints
	}

	s.breakpoints = list.New()

	for i := range bp {
		s.breakpoints.PushFront(&Breakpoint{value: bp[len(bp)-(i+1)]})
	}

	// for _, v := range bp {
	// 	err := s.tFindBreakpoints(sheet, v)
	// 	if err != nil {
	// 		return err
	// 	}
	// }

	return s.findBreakpoints(sheet, s.breakpoints.Len())
}

func (s *Sheet) findBreakpoints(sheet string, elCount int) error {
	rows, err := s.file.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("findBreakpoints: %v", err)
	}

	for _, row := range rows {
		if elCount <= 0 {
			break
		}
		for j, colCell := range row {
			bp := s.breakpoints.Front()
			if strings.Contains(colCell, bp.Value.(*Breakpoint).value) {
				bp.Value.(*Breakpoint).cellName, err = excelize.ColumnNumberToName(j + 1)
				if err != nil {
					return fmt.Errorf("findBreakpoints: %v", err)
				}
				// s.breakpoints.PushFront(s.breakpoints.Remove(bp))
				bp = bp.Next()
				elCount -= 1
			}
		}
	}
	return nil
}

func (s *Sheet) tFindBreakpoints(sheet string, breakpoint string) error {
	rows, err := s.file.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("findBreakpoints: %v", err)
	}

	for _, row := range rows {
		for j, colCell := range row {
			if strings.Contains(colCell, breakpoint) {
				fmt.Println(colCell, j)
				cn, err := excelize.ColumnNumberToName(j + 1)
				if err != nil {
					return fmt.Errorf("findBreakpoints: %v", err)
				}
				s.breakpoints.PushBack(&Breakpoint{value: breakpoint, cellName: cn})
				goto FINISH
			}
		}
	}
FINISH:
	return nil
}
