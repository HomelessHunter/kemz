package excelizer

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestFindBreakpoints(t *testing.T) {
	sheet, err := OpenSheet("рассчет.xlsx")
	require.NoError(t, err)
	defer func() {
		err = sheet.CloseSheet()
		require.NoError(t, err)
	}()

	err = sheet.SetBreakpoints("РВД", "№ п/п", "Поставщик", "Наименование", "июль", "август", "сентябрь", "Приход")
	require.NoError(t, err)
	sheet.ShowBreakpoints()
}
