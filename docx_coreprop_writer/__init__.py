#!/usr/bin/env python3

import datetime
import argparse
import yaml
import docx
from attrdict import AttrDict


def overwrite(meta, doc):
    doc_ = docx.Document(doc)  # type:docx.Document
    meta = meta.docx_coreprop
    if "author" in meta:
        """ author (unicode)
        Note: named ‘creator’ in spec.
        An entity primarily responsible for making the content of the resource. (Dublin Core)
        """
        doc_.core_properties.author = meta.author
    if "category" in meta:
        """ category (unicode)
        A categorization of the content of this package.
        Example values for this property might include: Resume, Letter, Financial Forecast, Proposal,
        Technical Presentation, and so on. (Open Packaging Conventions)
        """
        doc_.core_properties.category = meta.category
    if "comments" in meta:
        """comments (unicode)
        Note: named ‘description’ in spec.
        An explanation of the content of the resource.
        Values might include an abstract, table of contents, reference to a graphical representation
        of content, and a free-text account of the content. (Dublin Core)
        """
        doc_.core_properties.comments = meta.comments
    if "content_status" in meta:
        """content_status (unicode)
        The status of the content.
        Values might include “Draft”, “Reviewed”, and “Final”. (Open Packaging Conventions)
        """
        doc_.core_properties.content_status = meta.content_status
    if "created" in meta:
        """created (datetime)
        Date of creation of the resource. (Dublin Core)
        """
        doc_.core_properties.created = datetime.datetime.strptime(meta.created, "%Y-%m-%d")  # YYYY-MM-DD
    if "identifier" in meta:
        """identifier (unicode)
        An unambiguous reference to the resource within a given context. (Dublin Core)
        """
        doc_.core_properties.identifier = meta.identifier
    if "keywords" in meta:
        """keywords (unicode)
        A delimited set of keywords to support searching and indexing.
        This is typically a list of terms that are not available elsewhere
        in the properties. (Open Packaging Conventions)
        """
        doc_.core_properties.keywords = meta.keywords
    if "language" in meta:
        """language (unicode)
        The language of the intellectual content of the resource. (Dublin Core)
        """
        doc_.core_properties.language = meta.language
    if "last_modified_by" in meta:
        """last_modified_by (unicode)
        The user who performed the last modification. The identification is environment-specific.
        Examples include a name, email address, or employee ID.
        It is recommended that this value be as concise as possible. (Open Packaging Conventions)
        """
        doc_.core_properties.last_modified_by = meta.last_modified_by
    if "last_printed" in meta:
        """last_printed (datetime)
        The date and time of the last printing. (Open Packaging Conventions)
        """
        doc_.core_properties.last_printed = datetime.datetime.strptime(meta.last_printed, "%Y-%m-%d")
    if "modified" in meta:
        """modified (datetime)
        Date on which the resource was changed. (Dublin Core)
        """
        doc_.core_properties.modified = datetime.datetime.strptime(meta.modified, "%Y-%m-%d")
    if "revision" in meta:
        """revision (int)
        The revision number. This value might indicate the number of saves or revisions,
        provided the application updates it after each revision. (Open Packaging Conventions)
        """
        doc_.core_properties.revision = meta.revision
    if "subject" in meta:
        """subject (unicode)
        The topic of the content of the resource. (Dublin Core)
        """
        doc_.core_properties.subject = meta.subject
    if "title" in meta:
        """title (unicode)
        The name given to the resource. (Dublin Core)
        """
        doc_.core_properties.title = meta.title
    if "version" in meta:
        """version (unicode)
        The version designator.
        This value is set by the user or by the application. (Open Packaging Conventions)
        """
        doc_.core_properties.version = meta.version

    doc_.save(doc)


def main():
    parser = argparse.ArgumentParser(description="Reads yaml, overwrites DOCX core property")
    parser.add_argument("--input", "-I", help="yaml input filename")
    parser.add_argument("--output", "-O", help="docx output filename")

    args = parser.parse_args()
    with open(args.input, "r") as file:
        meta = AttrDict(yaml.load(file.read()))
    doc = args.output
    overwrite(meta, doc)


if __name__ == "__main__":
    main()
