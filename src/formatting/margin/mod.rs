mod bottom_margin;
mod left_margin;
mod right_margin;
mod top_margin;

pub use self::{bottom_margin::*, left_margin::*, right_margin::*, top_margin::*};

#[cfg(test)]
mod tests {
    use super::*;
    use hard_xml::XmlRead;

    #[test]
    fn decimal_widths_are_rounded() -> hard_xml::XmlResult<()> {
        assert_eq!(
            TopMargin::from_str(r#"<w:top w:w="100.0"/>"#)?.size,
            Some(100)
        );
        assert_eq!(
            BottomMargin::from_str(r#"<w:bottom w:w="100.4"/>"#)?.size,
            Some(100)
        );
        assert_eq!(
            LeftMargin::from_str(r#"<w:left w:w="100.5"/>"#)?.size,
            Some(101)
        );
        assert_eq!(
            RightMargin::from_str(r#"<w:right w:w="-100.5"/>"#)?.size,
            Some(-101)
        );
        Ok(())
    }
}
