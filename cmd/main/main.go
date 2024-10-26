/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)
   Copyright (c) 2024 mjanker4247

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

// Package main is a function demo
package main

import (
	"bytes"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/mjanker4247/go-docx"
)

// Returns a slice of image file paths from the specified directory
func getImageFiles(dir string) ([]string, error) {
	var images []string

	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, fmt.Errorf("failed to read directory: %w", err)
	}

	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		filePath := filepath.Join(dir, entry.Name())
		if isImage(filePath) {
			images = append(images, filePath)
		}
	}

	return images, nil
}

// Determines if a file is an image based on magic bytes
func isImage(filename string) bool {
	file, err := os.Open(filename)
	if err != nil {
		log.Printf("Failed to open file: %s", err)
		return false
	}
	defer file.Close()

	// Read the first 8 bytes to identify the image type
	magicBytes := make([]byte, 8)
	if _, err := file.Read(magicBytes); err != nil {
		log.Printf("Failed to read file: %s", err)
		return false
	}

	switch {
	case bytes.HasPrefix(magicBytes, []byte{0xFF, 0xD8, 0xFF}): // JPEG
		return true
	case bytes.HasPrefix(magicBytes, []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}): // PNG
		return true
	case bytes.HasPrefix(magicBytes, []byte{0x47, 0x49, 0x46, 0x38}): // GIF
		return true
	case bytes.HasPrefix(magicBytes, []byte{0x42, 0x4D}): // BMP
		return true
	case bytes.HasPrefix(magicBytes, []byte{0x49, 0x49}) || bytes.HasPrefix(magicBytes, []byte{0x4D, 0x4D}): // TIFF
		return true
	default:
		return false
	}
}

func main() {
	fileLocation := flag.String("f", "new-file.docx", "file location")
	dir := flag.String("d", "./images", "Directory containing images")
	flag.Parse()

	// Get a list of image files from the directory
	files, err := getImageFiles(*dir)
	if err != nil {
		log.Fatalf("Failed to get image files: %v", err)
	}

	var w *docx.Docx

	fmt.Printf("Preparing new document to write at %s\n", *fileLocation)

	w = docx.New().WithDefaultTheme().WithA4Page()

	// Add images with captions to the document
	for i, filePath := range files {
		caption := fmt.Sprintf("Image %d: %s", i+1, filepath.Base(filePath))

		// Add image to the document 
		p1 := w.AddParagraph().Justification("center")
		_, err = p1.AddInlineDrawingFrom(filePath)
		if err != nil {
			panic(err)
		}
		// Add a line break
		p2 := w.AddParagraph().Justification("center")
		p2.AddText(caption)
	}

	// add new paragraph
	// para1 := w.AddParagraph().Justification("distribute")

	// add text
	// para1.AddText("test").AddTab()
	// para1.AddText("size").Size("44").AddTab()
	// para1.AddText("color").Color("808080").AddTab()
	// para1.AddText("shade").Shade("clear", "auto", "E7E6E6").AddTab()
	// para1.AddText("bold").Bold().AddTab()
	// para1.AddText("italic").Italic().AddTab()
	// para1.AddText("underline").Underline("double").AddTab()
	// para1.AddText("highlight").Highlight("yellow").AddTab()
	// para1.AddText("font").Font("Consolas", "", "cs").AddTab()

	// para2 := w.AddParagraph().Justification("end")
	// para2.AddText("test all font attrs").
	// 	Size("44").Color("ff0000").Font("Consolas", "", "cs").
	// 	Shade("clear", "auto", "E7E6E6").
	// 	Bold().Italic().Underline("wave").
	// 	Highlight("yellow")

	// nextPara := w.AddParagraph()
	// nextPara.AddLink("google", `http://google.com`)

	// para3 := w.AddParagraph().Justification("center")
	// // add text
	// para3.AddText("一行2个 inline").Size("44")

	// para4 := w.AddParagraph().Justification("center")
	// r, err := para4.AddInlineDrawingFrom("testdata/fumiama.JPG")
	// if err != nil {
	// 	panic(err)
	// }
	// para4.AddTab().AddTab()
	// r, err = para4.AddInlineDrawingFrom("testdata/fumiama2x.webp")
	// if err != nil {
	// 	panic(err)
	// }

	// w.AddParagraph().AddPageBreaks()
	// para5 := w.AddParagraph().Justification("center")
	// // add text
	// para5.AddText("一行1个 横向 inline").Size("44")

	// para6 := w.AddParagraph()
	// _, err = para6.AddInlineDrawingFrom("testdata/fumiamayoko.png")
	// if err != nil {
	// 	panic(err)
	// }


	// p := w.AddParagraph().Justification("center")
	// p.AddText("测试 AutoShape w:ln").Size("44")
	// _ = p.AddAnchorShape(808355, 238760, "AutoShape", "auto", "straightConnector1",
	// 	&docx.ALine{
	// 		W:         9525,
	// 		SolidFill: &docx.ASolidFill{SrgbClr: &docx.ASrgbClr{Val: "000000"}},
	// 		Round:     &struct{}{},
	// 		HeadEnd:   &docx.AHeadEnd{},
	// 		TailEnd:   &docx.ATailEnd{},
	// 	},
	// )
	// _ = p.AddInlineShape(808355, 238760, "AutoShape", "auto", "straightConnector1",
	// 	&docx.ALine{
	// 		W:         9525,
	// 		SolidFill: &docx.ASolidFill{SrgbClr: &docx.ASrgbClr{Val: "000000"}},
	// 		Round:     &struct{}{},
	// 		HeadEnd:   &docx.AHeadEnd{},
	// 		TailEnd:   &docx.ATailEnd{},
	// 	},
	// )

	f, err := os.Create(*fileLocation)
	if err != nil {
		panic(err)
	}
	_, err = w.WriteTo(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
	fmt.Println("Document writen.")

	fmt.Println("End of main")
}
