#!/usr/bin/env python3

import datetime
import argparse
import yaml
import docx
from attrdict import AttrDict


def ensure_value(namespace, dest, default):
    """
    Thanks to https://stackoverflow.com/a/29335524/6592473
    """
    stored = getattr(namespace, dest, None)
    if stored is None:
        return default
    return stored


class store_dict(argparse.Action):
    """
        Thanks to https://stackoverflow.com/a/29335524/6592473
    """

    def __call__(self, parser, namespace, values, option_string=None):
        vals = dict(ensure_value(namespace, self.dest, {}))
        key, _, val = values.partition('=')
        vals[key] = val
        setattr(namespace, self.dest, vals)


def get_choice(extra, meta, key):
    """
    tries to get value of key from extra and returns value
    when fails try same from meta, else returns None
    :param dict extra:
    :param dict meta:
    :param str key:
    :return ret:
    """
    ret = extra.get(key, meta.get(key, None))
    return ret


def overwrite(meta, doc, extra):
    """
    :param dict meta:
    :param docx.Doc doc:
    :param dict extra:
    :return:
    """
    fn = doc
    doc = docx.Document(fn)  # type:docx.Document
    meta = AttrDict(meta.get("docx_coreprop"))
    extra = AttrDict(extra)
    if meta is not None:
        author = get_choice(extra, meta, "author")
        category = get_choice(extra, meta, "category")
        comments = get_choice(extra, meta, "comments")
        content_status = get_choice(extra, meta, "content_status")
        created = get_choice(extra, meta, "created")
        identifier = get_choice(extra, meta, "identifier")
        keywords = get_choice(extra, meta, "keywords")
        language = get_choice(extra, meta, "language")
        last_modified_by = get_choice(extra, meta, "last_modified_by")
        last_printed = get_choice(extra, meta, "last_printed")
        modified = get_choice(extra, meta, "modified")
        revision = get_choice(extra, meta, "revision")
        subject = get_choice(extra, meta, "subject")
        title = get_choice(extra, meta, "title")
        version = get_choice(extra, meta, "version")
        if author is not None:
            """ author (unicode)
            Note: named ‘creator’ in spec.
            An entity primarily responsible for making the content of the resource. (Dublin Core)
            """
            doc.core_properties.author = author
        if category is not None:
            """ category (unicode)
            A categorization of the content of this package.
            Example values for this property might include: Resume, Letter, Financial Forecast, Proposal,
            Technical Presentation, and so on. (Open Packaging Conventions)
            """
            doc.core_properties.category = category
        if comments is not None:
            """comments (unicode)
            Note: named ‘description’ in spec.
            An explanation of the content of the resource.
            Values might include an abstract, table of contents, reference to a graphical representation
            of content, and a free-text account of the content. (Dublin Core)
            """
            doc.core_properties.comments = comments
        if content_status is not None:
            """content_status (unicode)
            The status of the content.
            Values might include “Draft”, “Reviewed”, and “Final”. (Open Packaging Conventions)
            """
            doc.core_properties.content_status = content_status
        if created is not None:
            """created (datetime)
            Date of creation of the resource. (Dublin Core)
            """
            doc.core_properties.created = datetime.datetime.strptime(created, "%Y-%m-%d")  # YYYY-MM-DD
        if identifier is not None:
            """identifier (unicode)
            An unambiguous reference to the resource within a given context. (Dublin Core)
            """
            doc.core_properties.identifier = identifier
        if keywords is not None:
            """keywords (unicode)
            A delimited set of keywords to support searching and indexing.
            This is typically a list of terms that are not available elsewhere
            in the properties. (Open Packaging Conventions)
            """
            doc.core_properties.keywords = keywords
        if language is not None:
            """language (unicode)
            The language of the intellectual content of the resource. (Dublin Core)
            """
            doc.core_properties.language = language
        if last_modified_by is not None:
            """last_modified_by (unicode)
            The user who performed the last modification. The identification is environment-specific.
            Examples include a name, email address, or employee ID.
            It is recommended that this value be as concise as possible. (Open Packaging Conventions)
            """
            doc.core_properties.last_modified_by = last_modified_by
        if last_printed is not None:
            """last_printed (datetime)
            The date and time of the last printing. (Open Packaging Conventions)
            """
            doc.core_properties.last_printed = datetime.datetime.strptime(last_printed, "%Y-%m-%d")
        if modified is not None:
            """modified (datetime)
            Date on which the resource was changed. (Dublin Core)
            """
            doc.core_properties.modified = datetime.datetime.strptime(modified, "%Y-%m-%d")
        if revision is not None:
            """revision (int)
            The revision number. This value might indicate the number of saves or revisions,
            provided the application updates it after each revision. (Open Packaging Conventions)
            """
            doc.core_properties.revision = revision
        if subject is not None:
            """subject (unicode)
            The topic of the content of the resource. (Dublin Core)
            """
            doc.core_properties.subject = subject
        if title is not None:
            """title (unicode)
            The name given to the resource. (Dublin Core)
            """
            doc.core_properties.title = title
        if version is not None:
            """version (unicode)
            The version designator.
            This value is set by the user or by the application. (Open Packaging Conventions)
            """
            doc.core_properties.version = version

    doc.save(fn)


def main():
    parser = argparse.ArgumentParser(description="Reads yaml, overwrites DOCX core property")
    parser.add_argument("--input", "-I", help="yaml input filename", required=True)
    parser.add_argument("--output", "-O", help="docx output filename", required=True)
    parser.add_argument("--metadata", "-M", default={}, action=store_dict)

    args = parser.parse_args()
    with open(args.input, "r") as file:
        meta = dict(yaml.load(file.read()))
    doc = args.output
    extra = args.metadata
    # print(type(extra), type(meta))
    overwrite(meta, doc, extra)


if __name__ == "__main__":
    main()
