package main

import (
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Author struct to hold author details
type Author struct {
	Key  string `json:"key"`
	Name string `json:"name"`
}

type Subject struct {
	Name string `json:"name"`
	URL  string `json:"url"`
}

// BookInfo struct updated to use the Author struct
type BookInfo struct {
	ISBN       []string  `json:"isbn_13,omitempty"` // Changed to a slice
	Title      string    `json:"title"`
	Authors    []Author  `json:"authors"` // Correctly defined as a slice of Author
	Published  string    `json:"publish_date"`
	PageCount  int       `json:"number_of_pages"`
	Languages  []string  `json:"languages"` // Changed to a slice
	Categories []Subject `json:"subjects"`  // Changed to a slice of Subject
	ImageLinks struct {
		Thumbnail string `json:"large"`
	} `json:"cover"`
	Language string `json:"languages"`
}

func getBookInfoByISBN(isbn string) (*BookInfo, error) {
	isbn = strings.ReplaceAll(isbn, "-", "") // Clean ISBN
	isbn = strings.TrimSpace(isbn)
	url := fmt.Sprintf("https://openlibrary.org/api/books?bibkeys=ISBN:%s&format=json&jscmd=data", isbn)
	log.Printf("Fetching book info for ISBN: %s", isbn) // Log fetching process
	log.Printf("The URL is: %s", url)                   // Log fetching process

	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("failed to fetch book info: %s", resp.Status)
	}

	var bookData map[string]BookInfo
	if err := json.NewDecoder(resp.Body).Decode(&bookData); err != nil {
		return nil, err
	}

	log.Printf("Raw book data for ISBN: %s: %+v\n", isbn, bookData) // Log raw data

	bookInfo, exists := bookData["ISBN:"+isbn]
	if !exists {
		return nil, fmt.Errorf("no data found for ISBN: %s", isbn)
	}

	return &bookInfo, nil
}

func getBookInfoByTitleAuthor(title, author string) (*BookInfo, error) {
	title = strings.ReplaceAll(title, " ", "+")
	author = strings.ReplaceAll(author, " ", "+")
	url := fmt.Sprintf("https://openlibrary.org/search.json?title=%s&author=%s", title, author)
	log.Printf("The URL is: %s", url)
	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("failed to fetch book info for title %s and author %s", title, author)
	}

	var data struct {
		NumFound int        `json:"num_found"`
		Docs     []BookInfo `json:"docs"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&data); err != nil {
		return nil, err
	}

	if data.NumFound == 0 {
		return nil, nil
	}
	return &data.Docs[0], nil
}

func getBookInfoFromGoogleBooks(title, author string) (*BookInfo, error) {
	// Replace spaces with "+" for URL formatting
	title = strings.ReplaceAll(title, " ", "+")
	author = strings.ReplaceAll(author, " ", "+")
	url := fmt.Sprintf("https://www.googleapis.com/books/v1/volumes?q=intitle:%s+inauthor:%s", title, author)
	log.Printf("Fetching from Google Books API: %s", url)

	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	log.Printf("after response: %s", "tet")
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("failed to fetch book info from Google Books API: %s", resp.Status)
	}

	var googleData struct {
		TotalItems int `json:"totalItems"`
		Items      []struct {
			VolumeInfo BookInfo `json:"volumeInfo"`
		} `json:"items"`
	}

	if err := json.NewDecoder(resp.Body).Decode(&googleData); err != nil {
		return nil, err
	}

	if googleData.TotalItems == 0 {
		return nil, nil // No data found in Google Books either
	}

	return &googleData.Items[0].VolumeInfo, nil // Return the first item found
}

func enrichBookData(inputExcel, outputExcel string) {
	log.Println("Opening Excel file for reading:", inputExcel)
	f, err := excelize.OpenFile(inputExcel)
	if err != nil {
		log.Fatalf("failed to open Excel file: %v", err)
	}

	rows, err := f.GetRows("Book Sheet")
	if err != nil {
		log.Fatalf("failed to get rows: %v", err)
	}

	enrichedData := [][]string{}

	log.Println("Enriching book data...")
	for _, row := range rows[1:] { // Skip header row
		isbn := row[0]
		author := row[1]
		title := row[2]
		condition := row[3]

		var bookInfo *BookInfo
		var err error

		if isbn != "" {
			bookInfo, err = getBookInfoByISBN(isbn)
		}
		if err != nil || bookInfo == nil {
			bookInfo, err = getBookInfoByTitleAuthor(title, author)
		}
		if err != nil || bookInfo == nil {
			log.Printf("No data found for ISBN: %s, Title: '%s', Author: '%s', trying Google Books API...", isbn, title, author)
			bookInfo, err = getBookInfoFromGoogleBooks(title, author)
		}
		if err != nil || bookInfo == nil {
			log.Printf("No data found for ISBN: %s, Title: '%s', Author: '%s'", isbn, title, author) // Log when no data is found
			enrichedData = append(enrichedData, []string{
				isbn, author, title, condition, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A",
			})
			continue
		}

		enrichedData = append(enrichedData, []string{
			isbn,
			strings.Join(extractAuthorNames(bookInfo.Authors), ", "), // Extract author names
			bookInfo.Title,
			condition,
			bookInfo.Published,
			"N/A", // Placeholder for series
			fmt.Sprintf("%d", bookInfo.PageCount),
			bookInfo.Language,
			strings.Join(extractSubjectNames(bookInfo.Categories), ", "),
			bookInfo.ImageLinks.Thumbnail,
		})
	}

	// Check if output file exists and remove it if it does
	if _, err := os.Stat(outputExcel); err == nil {
		log.Printf("Output file %s already exists. Removing it...", outputExcel)
		if err := os.Remove(outputExcel); err != nil {
			log.Fatalf("failed to remove existing file: %v", err)
		}
	}

	log.Println("Creating output Excel file...")
	outputFile := excelize.NewFile()
	outputFile.NewSheet("Sheet1")

	// Write the header row
	header := []string{"ISBN", "author name", "book name", "book condition", "date of publication", "series", "page count", "language", "tags", "image links"}
	for col, value := range header {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1) // Start writing at row 1
		outputFile.SetCellValue("Sheet1", cell, value)
	}

	// Write the enriched data
	log.Println("Writing enriched data to Excel...")
	for rowIndex, data := range enrichedData {
		for colIndex, value := range data {
			cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+2) // Start writing at row 2
			outputFile.SetCellValue("Sheet1", cell, value)
		}
	}

	if err := outputFile.SaveAs(outputExcel); err != nil {
		log.Fatalf("failed to save enriched data: %v", err)
	}

	log.Printf("Enriched book data saved to %s\n", outputExcel)
}

// Function to extract author names from the Author struct
func extractAuthorNames(authors []Author) []string {
	names := make([]string, len(authors))
	for i, author := range authors {
		names[i] = author.Name
	}
	return names
}

// Function to extract subject names from the Subject struct
func extractSubjectNames(subjects []Subject) []string {
	names := make([]string, len(subjects))
	for i, subject := range subjects {
		names[i] = subject.Name
	}
	return names
}

func main() {
	log.SetFlags(log.Ldate | log.Ltime | log.Lshortfile) // Set log format
	enrichBookData("Books list.xlsx", "enriched_books.xlsx")
}
