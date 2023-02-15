package docxlib

import (
	"encoding/xml"
	"io"
)

type ParagraphChild struct {
	Link       *Hyperlink     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink,omitempty"`
	Run        *Run           `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r,omitempty"`
	Properties *RunProperties `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr,omitempty"`
}

type Paragraph struct {
	XMLName  xml.Name         `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	Children []ParagraphChild // Children will generate an unnecessary tag <Children> ... </Children> and we skip it by a self-defined xml.Marshaler

	file *Docx
}

func (p *Paragraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	err := e.EncodeToken(start)
	if err != nil {
		return err
	}
	for _, c := range p.Children {
		switch {
		case c.Link != nil:
			err = e.Encode(c.Link)
		case c.Run != nil:
			err = e.Encode(c.Run)
		case c.Properties != nil:
			err = e.Encode(c.Properties)
		default:
			continue
		}
		if err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	children := make([]ParagraphChild, 0, 64)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		switch tt := t.(type) {
		case xml.StartElement:
			var elem ParagraphChild
			switch tt.Name.Local {
			case "hyperlink":
				var value Hyperlink
				d.DecodeElement(&value, &start)
				id := getAtt(tt.Attr, "id")
				anchor := getAtt(tt.Attr, "anchor")
				if id != "" {
					value.ID = id
				}
				if anchor != "" {
					value.ID = anchor
				}
				elem.Link = &value
			case "r":
				var value Run
				d.DecodeElement(&value, &start)
				elem.Run = &value
				if value.InstrText == "" && value.Text == nil && value.Drawing == nil {
					continue
				}
			case "rPr":
				var value RunProperties
				d.DecodeElement(&value, &start)
				elem.Properties = &value
			default:
				continue
			}
			children = append(children, elem)
		}

	}
	p.Children = children
	return nil

}