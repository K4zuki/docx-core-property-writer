#!/usr/bin/env python3

from __future__ import print_function
import datetime
import sys
import argparse
import yaml
import docx
from attrdict import AttrDict
from docx_coreprop_writer.version import version

attr = ["author",
        "category",
        "comments",
        "content_status",
        "created",
        "identifier",
        "keywords",
        "language",
        "last_modified_by",
        "last_printed",
        "modified",
        "revision",
        "subject",
        "title",
        "version",
        ]


def ensure_value(namespace, dest, default):
    """ Thanks to https://stackoverflow.com/a/29335524/6592473 """

    stored = getattr(namespace, dest, None)
    if stored is None:
        return default
    return stored


class store_dict(argparse.Action):
    """ Thanks to https://stackoverflow.com/a/29335524/6592473 """

    def __call__(self, parser, namespace, values, option_string=None):
        vals = dict(ensure_value(namespace, self.dest, {}))
        key, _, val = values.partition('=')
        vals[key] = val
        setattr(namespace, self.dest, vals)


def get_choice(meta_ext, meta_file, key):
    """ tries to get meta_ext[key], then try meta_file[key]

    :param dict meta_ext:
    :param dict meta_file:
    :param str key:
    :return ret:
    """

    if meta_ext is not None:
        ret = meta_ext.get(key, None)
        if ret is None:
            ret = meta_file.get(key)
    else:
        ret = meta_file.get(key, None)
    return ret


def overwrite_meta(meta_file, filename, meta_ext):
    """ Overwrite DOCX core property from meta_file or meta_ext dictionaries
    When both dict has value for each for same key, meta_ext has priority

    :param dict meta_file:
    :param str filename:
    :param dict meta_ext:
    """
    if meta_file is not None:
        meta_file = meta_file.get("docx_coreprop")
    doc = docx.Document(filename)  # type:docx.Document

    meta = AttrDict({key: get_choice(meta_ext, meta_file, key) for key in attr})
    [print("{} = {}".format(key, val), file=sys.stderr) for key, val in meta.items()]
    if meta.author is not None:
        """ author (unicode)
        Note: named `creator` in spec.
        An entity primarily responsible for making the content of the resource. (Dublin Core)
        """
        doc.core_properties.author = meta.author
    if meta.category is not None:
        """ category (unicode)
        A categorization of the content of this package.
        Example values for this property might include: Resume, Letter, Financial Forecast, Proposal,
        Technical Presentation, and so on. (Open Packaging Conventions)
        """
        doc.core_properties.category = meta.category
    if meta.comments is not None:
        """comments (unicode)
        Note: named `description` in spec.
        An explanation of the content of the resource.
        Values might include an abstract, table of contents, reference to a graphical representation
        of content, and a free-text account of the content. (Dublin Core)
        """
        doc.core_properties.comments = meta.comments
    if meta.content_status is not None:
        """content_status (unicode)
        The status of the content.
        Values might include "Draft", "Reviewed", and "Final". (Open Packaging Conventions)
        """
        doc.core_properties.content_status = meta.content_status
    if meta.created is not None:
        """created (datetime)
        Date of creation of the resource. (Dublin Core)
        """
        doc.core_properties.created = datetime.datetime.strptime(meta.created, "%d-%b-%Y")  # DD-MMM-YYYY
    if meta.identifier is not None:
        """identifier (unicode)
        An unambiguous reference to the resource within a given context. (Dublin Core)
        """
        doc.core_properties.identifier = meta.identifier
    if meta.keywords is not None:
        """keywords (unicode)
        A delimited set of keywords to support searching and indexing.
        This is typically a list of terms that are not available elsewhere
        in the properties. (Open Packaging Conventions)
        """
        doc.core_properties.keywords = meta.keywords
    if meta.language is not None:
        """language (unicode)
        The language of the intellectual content of the resource. (Dublin Core)
        """
        doc.core_properties.language = meta.language
    if meta.last_modified_by is not None:
        """last_modified_by (unicode)
        The user who performed the last modification. The identification is environment-specific.
        Examples include a name, email address, or employee ID.
        It is recommended that this value be as concise as possible. (Open Packaging Conventions)
        """
        doc.core_properties.last_modified_by = meta.last_modified_by
    if meta.last_printed is not None:
        """last_printed (datetime)
        The date and time of the last printing. (Open Packaging Conventions)
        """
        doc.core_properties.last_printed = datetime.datetime.strptime(meta.last_printed, "%d-%b-%Y")
    if meta.modified is not None:
        """modified (datetime)
        Date on which the resource was changed. (Dublin Core)
        """
        doc.core_properties.modified = datetime.datetime.strptime(meta.modified, "%d-%b-%Y")
    if meta.revision is not None:
        """revision (int)
        The revision number. This value might indicate the number of saves or revisions,
        provided the application updates it after each revision. (Open Packaging Conventions)
        """
        doc.core_properties.revision = meta.revision
    if meta.subject is not None:
        """subject (unicode)
        The topic of the content of the resource. (Dublin Core)
        """
        doc.core_properties.subject = meta.subject
    if meta.title is not None:
        """title (unicode)
        The name given to the resource. (Dublin Core)
        """
        doc.core_properties.title = meta.title
    if meta.version is not None:
        """version (unicode)
        The version designator.
        This value is set by the user or by the application. (Open Packaging Conventions)
        """
        doc.core_properties.version = meta.version

    doc.save(filename)


def replace_style(meta_file, filename, style_ext):
    """
    :param dict meta_file:
    :param str filename:
    :param dict style_ext:
    :return:
    """

    doc = docx.Document(filename)  # type:docx.Document
    if meta_file is not None:
        meta_file = meta_file.get("docx_coreprop")
    para = get_choice(style_ext, meta_file, "paragraph")
    table = get_choice(style_ext, meta_file, "table")
    if para is not None:
        for key, val in para.items():
            for p in doc.paragraphs:
                if p.style.name == key:
                    print("{} -> {}".format(key, val), file=sys.stderr)
                    p.style = doc.styles[val]
    if table is not None:
        for key, val in table.items():
            for t in doc.tables:
                if t.style.name == key:
                    print("{} -> {}".format(key, val), file=sys.stderr)
                    t.style = doc.styles[val]
                    # print(t.style)

    # print(doc.styles)
    doc.save(filename)


def main():
    parser = argparse.ArgumentParser(description="Reads yaml, overwrites DOCX core property")
    parser.add_argument("--input", "-I", help="yaml input filename")
    parser.add_argument("--output", "-O", help="docx output filename")
    parser.add_argument("--metadata", "-M", default={}, action=store_dict)
    parser.add_argument("--paragraph", "-P", default=None, action=store_dict)
    parser.add_argument("--table", "-T", default=None, action=store_dict)
    parser.add_argument('--version', action='version', version=str(version))

    args = parser.parse_args()

    with open(args.input, "r") as file:
        meta_file = yaml.load(file.read(), Loader=yaml.SafeLoader)
    doc = args.output
    meta_ext = args.metadata
    style_ext = {"paragraph": args.paragraph, "table": args.table, }

    overwrite_meta(meta_file, doc, meta_ext)
    replace_style(meta_file, doc, style_ext)
    print("{} processed".format(doc), file=sys.stderr)


if __name__ == "__main__":
    main()
