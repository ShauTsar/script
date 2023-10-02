package main

import (
	"fmt"
	"github.com/gammban/numtow"
	"github.com/gammban/numtow/lang"
	"github.com/tealeg/xlsx"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/schema/soo/wml"
	"log"
	"strconv"
	"strings"
	"time"
)

//	func init() {
//		license.SetMeteredKey("e9d30fd832da56d79a0dc7339fb08d01f7a5a8e908ce4829cd714a3aad93ebc9")
//
// }
func init() {
	license.SetMeteredKey("e9d30fd832da56d79a0dc7339fb08d01f7a5a8e908ce4829cd714a3aad93ebc9")
	//unipdflicense.SetMeteredKey(os.Getenv(`bfd9f4dc7b86c4afc99ea457438885cbbb40532328f3a359e7a1bfa188d35d`))

}
func main() {
	license.GetLicenseKey()
	xlFile, err := xlsx.OpenFile("input.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	sliseOfSub := make(map[string]string)
	sliseOfRealise := make(map[string]string)
	sheet := xlFile.Sheets[0] // Выберите лист (например, первый лист)
	row := sheet.Rows[6]      // Выберите строку (например, первая строка)
	cell := row.Cells[1]
	rowDate := sheet.Rows[8]
	cellDate := rowDate.Cells[1]
	startRowRealise := 13
loop:
	for rowIndex := startRowRealise; rowIndex <= sheet.MaxRow; rowIndex++ {
		newRow := sheet.Rows[rowIndex]
		for columnIndex, newCell := range newRow.Cells {
			if columnIndex == 5 {
				value, err := newCell.FormattedValue()
				if err != nil {
					log.Println(err)
					continue
				}
				if value == "Итого по договорам:" {
					break loop

				}
				debt := newRow.Cells[7]
				debtValue, err := debt.FormattedValue()
				if err != nil {
					log.Println(err)
					continue
				}
				//fmt.Printf("%s\t", value)
				sliseOfRealise[debtValue] = value

			} else if columnIndex == 0 {
				value, err := newCell.FormattedValue()
				if err != nil {
					//log.Println(err)
					continue
				}
				if value == "Итого по договорам:" {
					break loop

				}
				debt := newRow.Cells[2]
				debtValue, err := debt.FormattedValue()
				if err != nil {
					//log.Println(err)
					continue
				}
				sliseOfSub[debtValue] = value
			}
		}
	}

	counteragent := cell.String()
	date := cellDate.String()
	fmt.Printf("Значение ячейки: %s\n", counteragent)
	fmt.Printf("Значение ячейки: %s\n", date)

	doc := document.New()
	para := doc.AddParagraph()
	para.AddRun().AddText(fmt.Sprintf("г. Екатеринбург                                                                                                                      \"%s\"", date))
	para.SetAlignment(wml.ST_JcRight)

	// Добавляем основной текст в документ
	mainText := `ЗАО «ЭнергоСтрой», в лице Генерального директора Бурнева Б.В., действующего на основании Устава, с одной стороны, и
  ИП Мерц Олеся Анатольевна, действующая на основании свидетельства о государственной регистрации физического лица ОГРНИП № 317554300059968, с другой стороны,  подписали настоящее Соглашение о нижеследующем:`

	para = doc.AddParagraph()
	para.AddRun().AddText(mainText)
	i := 1
	t, _ := time.Parse("01-02-06", date)
	dateString := t.Format("02.01.2006")
	delete(sliseOfRealise, "")
	for key, value := range sliseOfRealise {
		value = addY(value)
		forNds, money, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()
		para.AddRun().AddText(strconv.Itoa(i) + " " + counteragent + " по состоянию на " + dateString + "г. имеет задолженность перед ЗАО «ЭнергоСтрой» по " + value + ". в размере " + money + " руб" + ". (" + moneyText + " " + smallMoneyText + "), в т.ч. НДС " + ndsString + " руб. (" + rublesText + " " + kopecksText + " " + "). Срок исполнения обязательств наступил. Наличие указанной задолженности подтверждается Актом  сверки взаиморасчетов по состоянию на  " + dateString + "г.")
		i++
	}
	delete(sliseOfSub, "")
	for key, value := range sliseOfSub {
		value = addY(value)
		_, money, _ := niceType(key)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()
		para.AddRun().AddText(strconv.Itoa(i) + " ЗАО «ЭнергоСтрой» по состоянию на " + dateString + "г. имеет задолженность перед " + counteragent + " " + value + "." + "в размере " + money + " руб. (" + moneyText + " " + smallMoneyText + "), без НДС. Срок исполнения обязательств наступил.Наличие указанной задолженности подтверждается Актом  сверки взаиморасчетов по состоянию на  " + dateString + "г.")
		i++
	}
	para = doc.AddParagraph()
	para.AddRun().AddText(strconv.Itoa(i) + "	Стороны пришли к соглашению о зачете взаимных  требований в соответствии со ст. 410 ГК РФ по обязательствам, указанным в п. 1 - 9 настоящего соглашения, в размере 5 132 208,76 руб. (Пять миллионов сто тридцать две тысячи двести восемь рублей 76 копеек), в т.ч. НДС 855 368,13 руб. (Восемьсот пятьдесят пять тысяч триста шестьдесят восемь рублей 13 копеек")
	b := 1
	for key, value := range sliseOfRealise {
		value = addY(value)
		forNds, money, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()
		para.AddRun().AddText(strconv.Itoa(i) + "." + strconv.Itoa(b) + " Обязательства" + counteragent + " перед ЗАО «ЭнергоСтрой» по " + value + ". прекращаются в размере " + money + " руб" + ". (" + moneyText + " " + smallMoneyText + "), в т.ч. НДС " + ndsString + " руб. (" + rublesText + " " + kopecksText + " " + "). с " + dateString + "г.")
		b++
	}
	for key, value := range sliseOfSub {
		value = addY(value)
		_, money, _ := niceType(key)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()
		para.AddRun().AddText(strconv.Itoa(i) + "." + strconv.Itoa(b) + " Обязательства ЗАО «ЭнергоСтрой» перед " + counteragent + " по " + value + ". прекращаются в размере " + money + " руб" + ". (" + moneyText + " " + smallMoneyText + "),без НДС с " + dateString + "г.")
		b++
	}
	para = doc.AddParagraph()
	para.AddRun().AddText(strconv.Itoa(i) + " С момента подписания настоящего Соглашения стороны считают  себя свободными от обязательств, в размере, прекращенном зачетом согласно п.10 настоящего соглашения. \nНастоящее Соглашение составлено в 2-х подлинных экземплярах,  по одному для каждой из сторон. \nНастоящее Соглашение вступает в силу с момента его   подписания сторонами.")
	fileNameDate := time.Now().Format("02.01.2006")
	err2 := doc.SaveToFile(fileNameDate + ".docx")
	if err2 != nil {
		log.Fatalf("Ошибка при сохранении файла: %v", err2)
	}
}
func sumToText(text string) (string, string) {
	parts := strings.Split(text, ",")
	rublesPart := parts[0]
	kopecksPart := parts[1]
	rublesText := numtow.MustString(rublesPart, lang.RU) + " рублей"
	kopecksText := numtow.MustString(kopecksPart, lang.RU) + " копеек"
	return rublesText, kopecksText
}
func niceType(amountStr string) (float64, string, error) {
	// Убираем символ валюты и разделитель тысяч (если есть)
	amountStr = strings.ReplaceAll(amountStr, "₽", "")
	amountStr = strings.ReplaceAll(amountStr, " ", "")

	// Преобразовываем строку в число с плавающей точкой
	amount, err := strconv.ParseFloat(amountStr, 64)
	if err != nil {
		fmt.Println("Ошибка разбора суммы:", err)
		return 0, "", err
	}
	formattedAmount := strconv.FormatFloat(amount, 'f', 2, 64)
	formattedAmount = strings.ReplaceAll(formattedAmount, ".", ",")
	return amount, formattedAmount, nil
}
func addY(text string) string {
	text = strings.ToLower(text)
	parts := strings.Split(text, " ")

	// Ищем слово "договор" и добавляем "у" к нему
	for i, part := range parts {
		if part == "договор" {
			parts[i] = part + "у"
		} else if part == "года" {
			parts[i] = "г"
		}
	}

	// Объединяем подстроки обратно в одну строку
	text = strings.Join(parts, " ")

	return text
}

//func numberToText(number string) string {
//	number = strings.ReplaceAll(number, " ", "")
//	// Маппинг для текстового представления цифр
//	textMapping := map[string]string{
//		"0": "ноль",
//		"1": "один",
//		"2": "два",
//		"3": "три",
//		"4": "четыре",
//		"5": "пять",
//		"6": "шесть",
//		"7": "семь",
//		"8": "восемь",
//		"9": "девять",
//	}
//
//	// Маппинг для текстового представления десятков
//	tensMapping := map[string]string{
//		"10": "десять",
//		"11": "одиннадцать",
//		"12": "двенадцать",
//		"13": "тринадцать",
//		"14": "четырнадцать",
//		"15": "пятнадцать",
//		"16": "шестнадцать",
//		"17": "семнадцать",
//		"18": "восемнадцать",
//		"19": "девятнадцать",
//		"2":  "двадцать",
//		"3":  "тридцать",
//		"4":  "сорок",
//		"5":  "пятьдесят",
//		"6":  "шестьдесят",
//		"7":  "семьдесят",
//		"8":  "восемьдесят",
//		"9":  "девяносто",
//	}
//
//	// Маппинг для текстового представления разрядов
//	rankMapping := map[int]string{
//		3: "тысяч",
//		6: "миллионов",
//		9: "миллиардов",
//	}
//
//	// Разбиваем число на символы
//	digits := strings.Split(number, "")
//	length := len(digits)
//
//	var text string
//
//	for i, digit := range digits {
//		index := length - i
//
//		if digit == "0" {
//			continue // Пропускаем ноль
//		}
//
//		// Добавляем текст для разрядов (тысячи, миллионы и миллиарды)
//		if index > 1 && index%3 == 0 {
//			rank := rankMapping[index]
//			text += numberToText(digit) + " " + rank + " "
//		} else if index == 3 {
//			text += textMapping[digit] + " сотен "
//		} else if index == 2 {
//			if digit == "1" {
//				// Специальное представление для чисел 10-19
//				nextDigit := digits[i+1]
//				text += tensMapping[digit+nextDigit] + " "
//				break
//			} else {
//				text += tensMapping[digit] + " "
//			}
//		} else if index == 1 {
//			text += textMapping[digit] + " "
//		}
//	}
//
//	return text
//}
