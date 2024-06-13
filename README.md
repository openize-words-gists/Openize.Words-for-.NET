# Word Document SDK Usage Examples in C#

**Openize.Words.Examples (Formerly FileFormat.Words.Examples)** - C# code examples using [Openize.Words for .NET](https://github.com/openize-words/Openize-Words-for-.NET) (Formerly FileFormat.Words) - A robust native C# SDK for creating, loading, and modifying MS Word documents.

## Table of Contents
- [Create Word Document Paragraphs in C#](#create-word-document-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/b768640e5de61db628150a8b5bf5e923)
- [Read Word Document Paragraphs in C#](#read-word-document-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/b2eebe46d7445258d2b9ba612d90b362)
- [Modify Word Document Paragraphs in C#](#modify-word-document-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/501f3c7db4e5c015c493e784225a4ea9)
- [Create Word Document Images in C#](#create-word-document-images-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/fc8dd5d0ec414f7a1c0f4966cd0109a3)
- [Read Word Document Images in C#](#read-word-document-images-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/41892296f5a4554a85b45133358e4ea7)
- [Modify Word Document Images in C#](#modify-word-document-images-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/6a5e459efed49465aab84e33507148af)
- [Create Word Document Tables in C#](#create-word-document-tables-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/15390fa6c74c136ea2dbaf31fcea5f71)
- [Read Word Document Tables in C#](#read-word-document-tables-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/8763b0c8cdafdaa0178201c83d089a65)
- [Modify Word Document Tables in C#](#modify-word-document-tables-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/7f2d897e587120db72d4ea7f5e41b193)
- [Multiple Word Documents Concurrent Updating in C#](#multiple-word-documents-concurrent-updating-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/a39fe003185d962e6cd98071ed19bc08)
- [Create Word Paragraph Alignment in C#](#create-word-paragraph-alignment-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/dbf74b178de546aac7b5f9905ddd18e7)
- [Read Word Paragraph Alignment in C#](#read-word-paragraph-alignment-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/74de01e824f245f5b7759519a857a223)
- [Modify Word Paragraph Alignment in C#](#modify-word-paragraph-alignment-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/7a269bb5782142df6d7c358d5381e915)
- [Create Word Paragraph Indent in C#](#create-word-paragraph-indent-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/423cb924e480d848cc855649a596a51b)
- [Read Word Paragraph Indent in C#](#read-word-paragraph-indent-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/1752cd0a6ed9c9d9584629b8f74188ff)
- [Modify Word Paragraph Indent in C#](#modify-word-paragraph-indent-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/5e0292b9621d5b432688cdea1f5673d6)
- [Create Numbered Word Paragraphs in C#](#create-numbered-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/d4f95e410252b372496f05f452c50bc8)
- [Read Numbered Word Paragraphs in C#](#read-numbered-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/7022045cf26064df7388dbd883daa4cd)
- [Modify Numbered Word Paragraphs in C#](#modify-numbered-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/9b0128a2400c0d033d1d9bffe75519ee)
- [Create Roman Alphabetic Word Paragraphs in C#](#create-roman-alphabetic-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/e71983ac92406667e6b03eff8c142203)
- [Read Roman Alphabetic Word Paragraphs in C#](#read-roman-alphabetic-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/9bfed3dedf0fe2c085427a281cb9473e)
- [Modify Roman Alphabetic Word Paragraphs in C#](#modify-roman-alphabetic-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/0686c2de379c81c945c44f8192029621)
- [Create Multiple Multilevel List Paragraphs of Word Document in C#](#create-multiple-multilevel-list-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/5e903077f050411ab4ffc0848f838a34)
- [Read Multiple Multilevel List Paragraphs of Word Document in C#](#read-multiple-multilevel-list-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/cf5e6290218466bfc50f5feab50c2f85)
- [Modify Multiple Multilevel List Paragraphs of Word Document in C#](#modify-multiple-multilevel-list-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/a0f6f0413e309a5dbb27d7e366f6ffa5)
- [Create Multiple Frame Paragraphs of Word Document in C#](#create-multiple-frame-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/69335fe2b99e26584f57eef1b9c58468)
- [Read Multiple Frame Paragraphs of Word Document in C#](#read-multiple-frame-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/a577267cc39bae46ace887d0860422db)
- [Modify Multiple Frame Paragraphs of Word Document in C#](#modify-multiple-frame-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/openize-words-gists/08e214d1abbf7603657d9b42e3ed9822)
- [Resources](#resources)
- [System Requirements](#system-requirements)
- [Quick Start](#quick-start)
- [License](#license)
  
## [Create Word Document Paragraphs in C#](https://gist.github.com/openize-words-gists/b768640e5de61db628150a8b5bf5e923)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates paragraphs with heading styles defined by the Word document template.
* Adds normal paragraphs under each heading paragraph.
* Includes text runs with various fonts as per the template.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/b768640e5de61db628150a8b5bf5e923).        

## [Read Word Document Paragraphs in C#](https://gist.github.com/openize-words-gists/b2eebe46d7445258d2b9ba612d90b362)
* Loads an existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays associated styles as defined by the Word document template.
* Traverses through each run (text fragment) within each paragraph and displays fragment values.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/b2eebe46d7445258d2b9ba612d90b362)

## [Modify Word Document Paragraphs in C#](https://gist.github.com/openize-words-gists/501f3c7db4e5c015c493e784225a4ea9)
* Loads an existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Modifies paragraphs by prepending 'Modified Heading :' for styled paragraphs and 'Modified Run :' for each run within normal paragraphs, preserving the existing format.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/501f3c7db4e5c015c493e784225a4ea9)

## [Create Word Document Images in C#](https://gist.github.com/openize-words-gists/fc8dd5d0ec414f7a1c0f4966cd0109a3)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Loads images from the specified diretory and decodes using SkiaSharp.
* Creates a word document and appends loaded images to it.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/fc8dd5d0ec414f7a1c0f4966cd0109a3)

## [Read Word Document Images in C#](https://gist.github.com/openize-words-gists/41892296f5a4554a85b45133358e4ea7)
* Loads an existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses images and displays image metadata.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/41892296f5a4554a85b45133358e4ea7)

## [Modify Word Document Images in C#](https://gist.github.com/openize-words-gists/6a5e459efed49465aab84e33507148af)
* Loads an existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Gets images from the word document. Decodes image using SkiaSharp and encode to JPG.
* Resize image to 250(height) and 200(width).
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/6a5e459efed49465aab84e33507148af)

## [Create Word Document Tables in C#](https://gist.github.com/openize-words-gists/15390fa6c74c136ea2dbaf31fcea5f71)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates 5(rows) x 3(cols) tables with table styles defined by the Word document template.
* Appends each table to the body of the word document.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/15390fa6c74c136ea2dbaf31fcea5f71)      

## [Read Word Document Tables in C#](https://gist.github.com/openize-words-gists/8763b0c8cdafdaa0178201c83d089a65)
* Loads an existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses tables and displays associated styles as defined by the Word document template.
* Traverses through each row and then traverses columns within the row.
* Traverses through paragrpahs within each cell and displays paragraph plain text
* [Click here to explore gist](https://gist.github.com/openize-words-gists/8763b0c8cdafdaa0178201c83d089a65)

## [Modify Word Document Tables in C#](https://gist.github.com/openize-words-gists/7f2d897e587120db72d4ea7f5e41b193)
* Loads an existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Modifies tables by setting column widths to 2000
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/7f2d897e587120db72d4ea7f5e41b193)

## [Multiple Word Documents Concurrent Updating in C#](https://gist.github.com/openize-words-gists/a39fe003185d962e6cd98071ed19bc08)
* Loads 3 existing Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Appends paragraphs concurrently on three documents
* Saves the modified Word Documents.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/a39fe003185d962e6cd98071ed19bc08)

## [Create Word Paragraph Alignment in C#](https://gist.github.com/openize-words-gists/dbf74b178de546aac7b5f9905ddd18e7)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates paragraphs with different alignments including left, center, right and justify.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/dbf74b178de546aac7b5f9905ddd18e7)

## [Read Word Paragraph Alignment in C#](https://gist.github.com/openize-words-gists/74de01e824f245f5b7759519a857a223)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays its text along with alignment.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/74de01e824f245f5b7759519a857a223)

## [Modify Word Paragraph Alignment in C#](https://gist.github.com/openize-words-gists/7a269bb5782142df6d7c358d5381e915)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses through all paragraphs within the document.
* Modifies paragraphs by appending ' (alignment modified to justify)' with italic format and justify alignment.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/7a269bb5782142df6d7c358d5381e915)

## [Create Word Paragraph Indent in C#](https://gist.github.com/openize-words-gists/423cb924e480d848cc855649a596a51b)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates paragraphs with different indents including left, right, firstline and hanging.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/423cb924e480d848cc855649a596a51b)

## [Read Word Paragraph Indent in C#](https://gist.github.com/openize-words-gists/1752cd0a6ed9c9d9584629b8f74188ff)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays its text along with indentation.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/1752cd0a6ed9c9d9584629b8f74188ff)

## [Modify Word Paragraph Indent in C#](https://gist.github.com/openize-words-gists/5e0292b9621d5b432688cdea1f5673d6)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses through all paragraphs within the document.
* Modifies paragraphs by appending the relevant indent message in italic format and modifies the indent (if found) to 0.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/5e0292b9621d5b432688cdea1f5673d6)

## [Create Numbered Word Paragraphs in C#](https://gist.github.com/openize-words-gists/d4f95e410252b372496f05f452c50bc8)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates numbered paragraphs with nested levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/d4f95e410252b372496f05f452c50bc8)

## [Read Numbered Word Paragraphs in C#](https://gist.github.com/openize-words-gists/7022045cf26064df7388dbd883daa4cd)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays its text, numbering and level.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/7022045cf26064df7388dbd883daa4cd)

## [Modify Numbered Word Paragraphs in C#](https://gist.github.com/openize-words-gists/9b0128a2400c0d033d1d9bffe75519ee)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses through all paragraphs within the document.
* If numbered, modifies paragraphs by appending ' (numering removed)' with italic format and paragraph number is removed.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/9b0128a2400c0d033d1d9bffe75519ee)

## [Create Roman Alphabetic Word Paragraphs in C#](https://gist.github.com/openize-words-gists/e71983ac92406667e6b03eff8c142203)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates roman and alphabetic paragraphs with nested levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/e71983ac92406667e6b03eff8c142203)

## [Read Roman Alphabetic Word Paragraphs in C#](https://gist.github.com/openize-words-gists/9bfed3dedf0fe2c085427a281cb9473e)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays its text, roman/alphabetic status and level.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/9bfed3dedf0fe2c085427a281cb9473e)

## [Modify Roman Alphabetic Word Paragraphs in C#](https://gist.github.com/openize-words-gists/0686c2de379c81c945c44f8192029621)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses through all paragraphs within the document.
* If alphabetic, modifies paragraphs by appending ' (alphabetic removed)' with italic format and paragraph alphabetic is removed.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/0686c2de379c81c945c44f8192029621)

## [Create Multiple Multilevel List Paragraphs of Word Document in C#](https://gist.github.com/openize-words-gists/5e903077f050411ab4ffc0848f838a34)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates two multilevel lists with different prefixes at different levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/5e903077f050411ab4ffc0848f838a34)

## [Read Multiple Multilevel List Paragraphs of Word Document in C#](https://gist.github.com/openize-words-gists/cf5e6290218466bfc50f5feab50c2f85)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays its text, numbering id, numbering type and level.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/cf5e6290218466bfc50f5feab50c2f85)

## [Modify Multiple Multilevel List Paragraphs of Word Document in C#](https://gist.github.com/openize-words-gists/a0f6f0413e309a5dbb27d7e366f6ffa5)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses through all paragraphs within the document.
* If list paragraphs, modifies paragraphs by appending ' (numbering type changed to numeric)' with italic format and paragraph numbering type is changed to numeric.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/a0f6f0413e309a5dbb27d7e366f6ffa5)

## [Create Multiple Frame Paragraphs of Word Document in C#](https://gist.github.com/openize-words-gists/69335fe2b99e26584f57eef1b9c58468)
* Creates a new Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Generates two multilevel lists with different prefixes at different levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/69335fe2b99e26584f57eef1b9c58468)

## [Read Multiple Frame Paragraphs of Word Document in C#](https://gist.github.com/openize-words-gists/a577267cc39bae46ace887d0860422db)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses paragraphs and displays its text, numbering id, numbering type and level.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/a577267cc39bae46ace887d0860422db)

## [Modify Multiple Frame Paragraphs of Word Document in C#](https://gist.github.com/openize-words-gists/08e214d1abbf7603657d9b42e3ed9822)
* Loads a Word Document with structured content using [Openize.Words](https://www.nuget.org/packages/Openize.Words)
* Traverses through all paragraphs within the document.
* If list paragraphs, modifies paragraphs by appending ' (numbering type changed to numeric)' with italic format and paragraph numbering type is changed to numeric.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/openize-words-gists/08e214d1abbf7603657d9b42e3ed9822)


## Resources
* [Docs](https://www.openize.com)
* [API Reference](https://www.openize.com)
* [Articles](https://www.openize.com)

## System Requirements
* Make sure to have below nuget packages installed:
  * [Openize.Words for .NET](https://www.nuget.org/packages/Openize.Words)
  * [SkiaSharp](https://www.nuget.org/packages/SkiaSharp)

## Quick Start
```csharp
// Prepares directory at the root of your project.
// Default is 'Documents/Paragraph' unless specified otherwise as param.
var paragraphExamples = new ParagraphExamples();
// Creates a word document with paragraphs and saves word document to the specified 
// directory. Default is 'Documents/Paragraph/WordParagraphs.docx' unless specified otherwise as param.
paragraphExamples.CreateWordParagraphs();
// Reads Paragraphs from the specified Word Document and displays plain text and formatting.
// Default is 'Documents/Paragraph/WordParagraphs.docx' unless specified otherwise as param.
paragraphExamples.ReadWordParagraphs();
// Modifies Paragraphs in the specified Word Document and saves the modified word document.
// Default document to modify is 'Documents/Paragraph/WordParagraphs.docx' unless specified otherise as param.
// Default modified document is saved as 'Documents/Paragraph/WordParagraphs.docx' unless specified otherise as param.
paragraphExamples.ModifyWordParagraphs();
```

## License

This project is licensed under the MIT License - see the [LICENSE]([LICENSE](https://github.com/openize-words-gists/Openize.Words-for-.NET/blob/main/LICENSE)) file for details.
