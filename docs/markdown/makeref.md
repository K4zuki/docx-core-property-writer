\toc

\newpage

::: {custom-style="Instruction"}
Instruction text
:::

# \<Heading Unnumbered\> {-}

## \<Heading Unnumbered 2\> {-}

### \<Heading Unnumbered 3\> {-}

#### \<Heading Unnumbered 4\> {-}

\newpage

# \<Heading 1\>

## \<Heading 2\>

### \<Heading 3\>

#### \<Heading 4\>

##### \<Heading 5\>

\<Body\>

<!--
```{=openxml}
<w:p>
<w:pPr>
<w:sectPr>
<w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:pPr>
</w:p>
```
-->

Table: \<Table Caption\> {#tbl:table-caption}

| \<Table Head Left\> | \<Table Head Centre\> | Head | Head | Head | Head | Head |
|:--------------------|:---------------------:|:-----|:-----|:-----|:-----|:-----|
| Cell                |         Cell          | Cell | Cell | Cell | Cell | Cell |

<!--
::: {.table width=[0.2,0.2,0.2,0.1,0.1,0.1,0.1]}
:::
-->

::: {custom-style="Note"}
\<Note\>
:::

::: {custom-style="Graphic Anchor"}
\<Insert graphic here (Graphic Anchor)\>
:::

::: {custom-style="Figure Caption"}
Figure 1: \<Figure Caption\>
:::

Numbered equation:

::: {.table width=[1.0] noheader=true}

| $$Type~equation~here$${#eq:equation-1} |
|:---------------------------------------|
|                                        |

:::
\

:::{custom-style="Bullet List 1"}
\<Bullet List 1\>
:::

:::{custom-style="Bullet List 2"}
\<Bullet List 2\>
:::
\

:::{custom-style="Feature List 1"}
\<Feature List 1\>
:::

:::{custom-style="Feature List 2"}
\<Feature List 2\>
:::
\

:::{custom-style="Numbered List 1"}
\<Numbered List 1\>
:::

:::{custom-style="Numbered List 2"}
\<Numbered List 2\>
:::

:::{custom-style="Numbered List 3"}
\<Numbered List 3\>
:::
\

:::{custom-style="Reference List"}
\<Reference List\>
:::
\

\newpage

::: {custom-style="Appendix Heading 1"}
\<Appendix Title\>
:::

::: {custom-style="Appendix Heading 2"}
\<Appendix Subsection\>
:::

::: {custom-style="Appendix Heading 3"}
\<Appendix Heading 3\>
:::

::: {custom-style="Appendix Heading 4"}
\<Appendix Heading 4\>
:::

::: {custom-style="Appendix Heading 5"}
\<Appendix Heading 5\>
:::

```bash
$ docker run --rm -it -v $PWD:/workdir k4zuki/pandocker-alpine
$ pandoc -o reference.docx --print-default-data-file reference.docx
$ pandoc -o custom-reference.docx --highlight-style=kate reference.docx
```

\

```python
import pprint
import docx

d = docx.Document("custom-reference.docx")
pprint.pprint(sorted([s.name for s in d.styles if "Tok" in s.name]))
```

\

:::{custom-style="Source Code"}
[Verbatim Char]{custom-style="Verbatim Char"}\
[AlertTok]{custom-style="AlertTok"}\
[AnnotationTok]{custom-style="AnnotationTok"}\
[AttributeTok]{custom-style="AttributeTok"}\
[BaseNTok]{custom-style="BaseNTok"}\
[BuiltinTok]{custom-style="BuiltinTok"}\
[CharTok]{custom-style="CharTok"}\
[CommentTok]{custom-style="CommentTok"}\
[CommentVarTok]{custom-style="CommentVarTok"}\
[ConstantTok]{custom-style="ConstantTok"}\
[ControlFlowTok]{custom-style="ControlFlowTok"}\
[DataTypeTok]{custom-style="DataTypeTok"}\
[DecValTok]{custom-style="DecValTok"}\
[DocumentationTok]{custom-style="DocumentationTok"}\
[ErrorTok]{custom-style="ErrorTok"}\
[ExtensionTok]{custom-style="ExtensionTok"}\
[FloatTok]{custom-style="FloatTok"}\
[FunctionTok]{custom-style="FunctionTok"}\
[ImportTok]{custom-style="ImportTok"}\
[InformationTok]{custom-style="InformationTok"}\
[KeywordTok]{custom-style="KeywordTok"}\
[NormalTok]{custom-style="NormalTok"}\
[OperatorTok]{custom-style="OperatorTok"}\
[OtherTok]{custom-style="OtherTok"}\
[PreprocessorTok]{custom-style="PreprocessorTok"}\
[RegionMarkerTok]{custom-style="RegionMarkerTok"}\
[SpecialCharTok]{custom-style="SpecialCharTok"}\
[SpecialStringTok]{custom-style="SpecialStringTok"}\
[StringTok]{custom-style="StringTok"}\
[VariableTok]{custom-style="VariableTok"}\
[VerbatimStringTok]{custom-style="VerbatimStringTok"}\
[WarningTok]{custom-style="WarningTok"}\

This is a pen
:::
\

```markdown
Source Code
```

\newpage

[`docs/Makefile`](./Makefile){.listingtable type=makefile #lst:docs-makefile}

[`Makefile` (from line 10)](../Makefile){.listingtable from=10 type=makefile #lst:makefile}

[`setup.py` (to line 50)](../setup.py){.listingtable to=50 type=python}

\newpage

[Link](https://google.com)

> Block Quote\
> Block Quote

Table: Sample Table {#tbl:sample-table}

| This |   is   | Table |
|:-----|:------:|------:|
| Left | Center | Right |

<!--
::: {.table width=[0.3,0.4,0.3]}
:::
-->

\newpage

::: {#fig:images}

::: {.table noheader=true width=[1.0]}

| ![Front image 1](images/front-image.png){#fig:front-image-1 width=130mm} |
|:------------------------------------------------------------------------:|
|                                                                          |

:::
::: {.table noheader=true width=[0.5,0.5]}

| ![Front image 2](images/front-image.png){#fig:front-image-2 width=70mm} | ![Front image 3](images/front-image.png){#fig:front-image-3 width=70mm} |
|:-----------------------------------------------------------------------:|:-----------------------------------------------------------------------:|
|                                                                         |                                                                         |

:::

Only works with DOCX output
:::

![Front image](images/front-image.png){#fig:front-image}
