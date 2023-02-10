package docxlib

// Color allows to set run color
func (r *Run) Color(color string) *Run {
	r.RunProperties.Color = &Color{
		Val: color,
	}

	return r
}

// Size allows to set run size
func (r *Run) Size(size string) *Run {
	r.RunProperties.Size = &Size{
		Val: size,
	}

	return r
}
