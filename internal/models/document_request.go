package models

import "fmt"

// DocumentRequest is the new v2 contract.
// Supports JSON and XML.
//
// Example JSON:
// {
//   "format": "excel",
//   "templateId": "default",
//   "layout": { "showPageNum": true },
//   "data": [ {"name":"Ana","age":30}, {"name":"Bruno"} ]
// }
//
// Example XML:
// <DocumentRequest>
//   <Format>pdf</Format>
//   <TemplateId></TemplateId>
//   <Layout>
//     <ShowPageNum>true</ShowPageNum>
//   </Layout>
//   <Data>
//     <Item><name>Ana</name><age>30</age></Item>
//     <Item><name>Bruno</name></Item>
//   </Data>
// </DocumentRequest>
type DocumentRequest struct {
	Format     DocumentFormatEnum `json:"format" xml:"Format" example:"excel"`
	TemplateID string             `json:"templateId,omitempty" xml:"TemplateId,omitempty" example:"default"`
	Layout     *LayoutConfig      `json:"layout,omitempty" xml:"Layout,omitempty"`
	// Data is flexible: it can be an array/object (single dataset) or an object of datasets for Excel sheets.
	Data DataPayload `json:"data" xml:"Data" swaggertype:"object"`
}

func (r *DocumentRequest) Validate() error {
	if !r.Format.IsValid() {
		return fmt.Errorf("format is required")
	}
	// Normalize legacy fields for backward compatibility.
	if r.Layout != nil && r.Layout.ShowPageNum != nil {
		if r.Layout.Footer == nil {
			r.Layout.Footer = &FooterConfig{}
		}
		if r.Layout.Footer.PageNumber == nil {
			r.Layout.Footer.PageNumber = &PageNumberConfig{}
		}
		if *r.Layout.ShowPageNum {
			r.Layout.Footer.PageNumber.Enabled = true
			if r.Layout.Footer.Show == nil {
				v := true
				r.Layout.Footer.Show = &v
			}
		}
	}
	if err := r.Layout.Validate(); err != nil {
		return err
	}
	if r.Layout != nil && len(r.Layout.Sheets) > 0 {
		for i := range r.Layout.Sheets {
			ds := r.Data.Get(r.Layout.Sheets[i].DataSource)
			if ds.IsEmpty() {
				name := r.Layout.Sheets[i].DataSource
				if name == "" {
					name = "default"
				}
				return fmt.Errorf("dataSource '%s' (sheet '%s') is empty or missing", name, r.Layout.Sheets[i].Name)
			}
		}
	}
	if r.Data.IsEmpty() {
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	return nil
}
