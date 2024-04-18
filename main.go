package main

import (
	"fmt"
	"net/http"
	"sync"

	"github.com/xuri/excelize/v2"
	"gopkg.in/gomail.v2"
)

func main() {

}

type Pegawai struct {
	NPP         string
	NIP         string
	Nama        string
	KDW         string
	Unit        string
	UnitBesaran string
	Email       string
}

func SendNotif(w http.ResponseWriter, r *http.Request) {
	f, err := excelize.OpenFile("data.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	// runtime.GOMAXPROCS(2)
	var wg sync.WaitGroup

	for i, row := range rows {
		if i == 0 {
			continue
		} else if i == 10 {
			break
		}

		if len(row) < 7 {
			continue
		}

		pegawai := Pegawai{
			NPP:         row[0],
			NIP:         row[1],
			Nama:        row[2],
			KDW:         row[3],
			Unit:        row[4],
			UnitBesaran: row[5],
			Email:       row[6],
		}

		wg.Add(1)
		go func(email string) {
			defer wg.Done()
			err := sendEmail(email)
			if err != nil {
				fmt.Println(err)
			}
		}(pegawai.Email)
		// fmt.Sprintf("%v: %v", i, pegawai.Email)

	}

	wg.Wait()

}

func sendEmail(email string) error {
	sender := ""
	smtpHost := ""
	smtpPort := 0
	username := ""
	password := ""

	m := gomail.NewMessage()

	m.SetHeader("From", sender)
	m.SetHeader("To", email)

	m.SetHeader("Subject", "Data from Excel!")
	m.SetBody("text/plain", "test send Email!")

	d := gomail.NewDialer(smtpHost, smtpPort, username, password)

	if err := d.DialAndSend(m); err != nil {
		return err
	} else {
		return nil
	}
}
