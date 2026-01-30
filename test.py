import json
from wordhelpers import (
    build_table,
    replace_placeholder_with_table,
    get_para_by_string,
    insert_paragraph_after,
    insert_text_into_row,
    insert_text_by_table_coords,
    generate_table,
)
from docx import Document

word_template: str = "word_template.docx"
output_docx: str = "output.docx"

if __name__ == "__main__":
    # Create a new Word document
    doc = Document(word_template)

    # Define a sample table dictionary
    table_dict = {
        "style": "test_tbl_style",
        "rows": [
            {
                "cells": [
                    {
                        "background_color": "#4BFF33",
                        "paragraphs": [
                            {
                                "style": None,
                                "alignment": "center",
                                "text": ["Centered (green back)"],
                            }
                        ],
                    },
                    {
                        "background_color": None,
                        "paragraphs": [
                            {"style": None, "text": ["No alignment (left)"]}
                        ],
                    },
                    {
                        "background_color": None,
                        "paragraphs": [
                            {
                                "style": None,
                                "alignment": "justify",
                                "text": ["Justify Aligned Text"],
                            }
                        ],
                    },
                ]
            },
            {
                "cells": [
                    {
                        "background_color": "#4BFF33",
                        "paragraphs": [
                            {
                                "style": None,
                                "alignment": "center",
                                "text": ["Centered Text, green back, 1 merge"],
                            }
                        ],
                    },
                    "merge",
                    {
                        "background_color": None,
                        "paragraphs": [{"text": ["No alignment (left)"]}],
                    },
                ]
            },
            {
                "cells": [
                    {
                        "background_color": "#4BFF33",
                        "paragraphs": [
                            {
                                "style": "test_bold",
                                "alignment": "right",
                                "text": [
                                    "Right-Aligned Bold Multi-Line Text",
                                    "green background",
                                    "2 merges",
                                ],
                            }
                        ],
                    },
                    "merge",
                    "merge",
                ]
            },
            {
                "cells": [
                    {
                        "table": {
                            "style": "test_tbl_style",
                            "rows": [
                                {
                                    "cells": [
                                        {
                                            "background_color": "#FF5733",
                                            "paragraphs": [
                                                {
                                                    "style": None,
                                                    "alignment": "right",
                                                    "text": ["Nested"],
                                                }
                                            ],
                                        },
                                        {
                                            "background_color": None,
                                            "paragraphs": [
                                                {
                                                    "style": None,
                                                    "alignment": "left",
                                                    "text": ["Table"],
                                                }
                                            ],
                                        },
                                    ]
                                }
                            ],
                        }
                    },
                    "merge",
                    "merge",
                ]
            },
        ],
    }

    table_dict["rows"].append(
        insert_text_into_row(
            [
                "auto row cell 1",
                "auto row cell 2 merged",
                "merge",
            ]
        )
    )

    table_dict = insert_text_by_table_coords(
        table_dict, 4, 0, "Inserted text into (4,0)"
    )

    # Build the table and add it to the document
    table = build_table(doc, table_dict)
    replace_placeholder_with_table(doc, "{{heading1_table}}", table)

    table = build_table(doc, table_dict, remove_leading_para=False)
    replace_placeholder_with_table(doc, "{{heading2_table}}", table)

    para = get_para_by_string(doc, "{{heading3_table}}")
    insert_paragraph_after(
        para, "This is a new paragraph inserted after heading3_table placeholder."
    )

    test_table: dict = generate_table(
        num_rows=2,
        num_cols=3,
        header_row=["Header 1", "Header 2", "Header 3"],
        style="test_tbl_style",
    )
    table = build_table(doc, test_table)
    replace_placeholder_with_table(doc, "{{heading4_table}}", table)

    # Save the document
    doc.save(output_docx)
