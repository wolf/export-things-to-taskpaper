import io
import sys
from operator import itemgetter
from pathlib import Path
from typing import Any, Iterable, List, NewType, Optional, TextIO
from typing_extensions import Annotated

import pyperclip  # type: ignore
import things  # type: ignore
import typer


__all__ = [  # unfortunately, this only applies where "from export_things_to_taskpaper import *"
    "export",
    "get_all_projects_for_area",
    "get_all_projects_with_no_area",
    "get_all_todos_for_project_in_order",
    "get_all_todos_for_area",
    "get_all_todos_with_no_area",
    "list_to_uuids",
    "print_uuids",
    "uuid_to_item",
    "write_all_items_in_area",
    "write_project",
    "write_projects",
    "write_todo",
    "write_todos",
]


UUID = NewType("UUID", str)


class UUIDDoesNotResolveError(Exception):
    def __init__(self, uuid: str, type: str | None) -> None:
        type_message = f", or else is not of type '{type}'" if type else ""
        super().__init__(f"The UUID {uuid} is not in the database{type_message}")


# Note: for projects to import correctly, in OmniFocus, you must import to "Projects", not to the "Inbox".

# Note that to reduce the total number of queries, and to avoid problems in how the things.py library works, anytime we call `things.tasks`, we pass `include_items=True` so we get everything underneath it, too.


def uuid_to_item(uuid: UUID, type: str | None = None) -> dict[str, Any]:
    """
    Return the dictionary that represents the item with the given UUID

    If uuid is not in the database, or if uuid _is_ in the database, but isn't the type asked for,
    raise a UUIDDoesNotResolveError.

    Only interactive users use UUIDs.  Only interactive users end up using this function.
    """
    item = None
    try:
        if type != "area":  # area isn't really a type; caller doesn't have to say this, but it makes sense
            item = things.tasks(uuid=uuid, type=type, include_items=True)
        if type == "area" or not item:
            item = things.areas(uuid=uuid, include_items=True)
    except ValueError:
        raise UUIDDoesNotResolveError(uuid, type)
    if not item:
        raise UUIDDoesNotResolveError(uuid, type)
    return item


def _resolve_item_if_needed(item: UUID | dict[str, Any], type: str | None = None) -> dict[str, Any]:
    """
    Return the dictionary that represents an item, whether by looking up the UUID, or just noticing you already had the item

    When run as a script, we never have to resolve UUIDs; but when used interactively the user may call any of the
    available public function with a UUID instead of a dictionary.  The functions are all written to work witth dicts
    so the first thing they do is get the dict if they didn't already have it.
    """
    if isinstance(item, str):
        # UUID just means str, so if it _is_ a string, let's get the actual item
        return uuid_to_item(item, type)
    # otherwise, it doesn't need to be resolved
    return item


def list_to_uuids(items: list[dict[str, Any]]) -> list[UUID]:
    """
    Given a list of items as represented by dicts, extract their UUIDs and return the entire collection as a list

    If you need a set, it's easy to get: set(list_to_uuids(items)).  If we returned a set, you couldn't get
    back to the list you started with because you would have lost order and duplicates.

    Only interactive users use UUIDs.  Only interactive users end up using this function.
    """
    return [item["uuid"] for item in items]


def print_uuids(uuids: Iterable[UUID]) -> None:
    """
    Print to stdout a comma separated collection of UUIDs, one per line

    Only interactive users use UUIDs.  Only interactive users end up using this function.
    """
    print(",\n".join(uuids))


def _write_note_if_any(item: dict[str, Any], stream: TextIO, depth: str) -> None:
    """
    Look for a note associated with item, and if there is one, write it to stream

    Because this function is not to be called interactively, it needn't accept a UUID or have defaults.
    The entire note is written out indented to depth, which is a string of zero or more tabs.
    """
    if note := item.get("notes"):
        stream.write(f"{depth}")  # make sure to indent the very first line of the note
        stream.write(f"\n{depth}".join(note.splitlines()))  # if there's more than one line, indent them all
        stream.write("\n")


def _write_checklist_if_any(todo: dict[str, Any], stream: TextIO, depth: str) -> None:
    """
    Look for a checklist associated with todo, and if there is one, write it to stream

    Only todos have checklists. The entire checklist is written out indented to depth, which is a string of zero or more tabs.

    Because this function is not to be called interactively, it needn't accept a UUID or have defaults.
    """
    if checklist := todo.get("checklist", []):
        for checklist_item in checklist:
            stream.write(f"""{depth}- {checklist_item["title"]}\n""")


def _get_omnifocus_parameters(item: dict[str, Any], area_name: str | None = None) -> str:
    """
    Return a string with application specific properties that will be tacked on to projects and todos

    This is a special list of elements of the form @key(value).  The exact list produced here are the
    ones that have special meaning to OmniFocus.
    """
    result = ""
    if when := item.get("start_date"):
        result += f" @defer({when})"
    if when := item.get("deadline"):
        result += f" @due({when})"
    tags = item.get("tags", [])  # yes, we'll tell OmniFocus any of the actual tags this item has
    if start := item.get("start"):
        tags += [start]  # ...but we'll also add "Anytime", "Someday", or "Inbox"
    if area_name:
        area_name = f"area-{area_name.replace(" ", "-")}"
        tags += [
            area_name
        ]  # ...and if this item was at the top-level of an area, make a tag for that (so it's easy to move by hand)
    result += f" @tags({", ".join(sorted(set(tags)))})"  # set to eliminate duplicates, sorted just because I want to

    # TODO: handle repeating items.  Ugh.  This looks annoying.  Note that things.py, the library I'm
    #   using for all this just doesn't handle repeating items.  I can't even see if a given item _is_
    #   repeating let alone any details.  If I could, then at least I could mark the item with a tag
    #   for fixup.  I think I'm going to pass for now.

    return result


def write_todo(
    todo: UUID | dict[str, Any],
    stream: TextIO = sys.stdout,
    depth: str = "",
    area_name: str | None = None,
) -> None:
    """
    Write one todo to stream

    The entire todo is indented to depth.  If the todo has a note and/or a checklist, they are also written out,
    indented one tab further than the todo itself.  Because this function might also be called by an interactive
    user, it accepts a UUID, and has defaults for all other parameters suitable for interactive use.
    """
    resolved_todo = _resolve_item_if_needed(todo, type="to-do")  # save todo in case it's a UUID and can't be resolved
    if not resolved_todo:
        assert isinstance(todo, str)
        raise UUIDDoesNotResolveError(
            todo, "to-do"
        )  # this can only happen for an interactive user, because only interactive users use UUIDs

    stream.write(
        f"""{depth}- {resolved_todo["title"]}{_get_omnifocus_parameters(resolved_todo, area_name=area_name)}\n"""
    )
    _write_note_if_any(resolved_todo, stream, depth + "\t")
    _write_checklist_if_any(resolved_todo, stream, depth + "\t")


def write_todos(
    todos: list[UUID] | list[dict[str, Any]],
    stream: TextIO = sys.stdout,
    depth: str = "",
    area_name: str | None = None,
) -> None:
    """
    Write a list of todos, one at a time, to stream

    The todos are all indented to the same depth.  Because this function might also be called by an interactive
    user, it accepts a list of UUIDs, and has defaults for all other parameters.  It does its work by calling write_todo
    which can take a UUID or a dict, so no testing has to happen in the loop.
    """
    for todo in todos:
        write_todo(todo, stream, depth, area_name)


def get_all_todos_for_project_in_order(
    project: UUID | dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Extract all the todos from a single project and return them as a list

    The return value is a list of dictionaries.  The list is sorted by index, so the it should be in the order
    the todos actually appear in the project.  Interactive users might want the result to be a list of UUIDs.
    That's easy to get: list_to_uuids(get_all_todos_for_project_in_order(project)).  Because this function might
    also be called by an interactive user, it accepts a project UUID.

    This function is written with the end goal of importing into OmniFocus in mind.  In OmniFocus, projects
    don't have headings.  In fact, I _think_ you can't express them in TaskPaper either.  In Things, when a
    project has headings, the todos under those headings are children of the heading object, not of the project.

    If there are no todos anywhere under this project, this function returns an empty list.
    """
    resolved_project = _resolve_item_if_needed(project, "project")
    if not resolved_project:
        assert isinstance(project, str)
        raise UUIDDoesNotResolveError(
            project, "project"
        )  # this can only happen for an interactive user, because only interactive users use UUIDs

    project_items = resolved_project.get("items", [])

    todos_directly_under_the_project = [todo for todo in project_items if todo.get("type") == "to-do"]

    headings = sorted(
        [heading for heading in project_items if heading.get("type") == "heading"],
        key=itemgetter("index"),
    )

    todos_under_headings = []
    for heading in headings:
        todos_under_headings += sorted(
            [todo for todo in heading.get("items", []) if todo.get("type") == "to-do"],
            key=itemgetter("index"),
        )

    return todos_directly_under_the_project + todos_under_headings


def write_project(project: UUID | dict[str, Any], stream: TextIO = sys.stdout, depth: str = "") -> None:
    """
    Write a single project (and all its todos) to stream

    Can't express headings in TaskPaper or OmniFocus, so we don't write or have any functions to write headings.
    The project is indented to depth.  Its note, if any, and all its todos are indented one tab-stop further.  Because this function
    might also be called by an interactive user, it accepts a UUID for the project, and defaults for all the
    other parameters.
    """
    resolved_project = _resolve_item_if_needed(project, "project")
    if not resolved_project:
        assert isinstance(project, str)
        raise UUIDDoesNotResolveError(
            project, "project"
        )  # this can only happen for an interactive user, because only interactive users use UUIDs

    # First write the project entry itself
    area_name = resolved_project.get("area_title")
    stream.write(
        f"""{depth}{resolved_project["title"]}:{_get_omnifocus_parameters(resolved_project, area_name=area_name)}\n"""
    )
    _write_note_if_any(resolved_project, stream, depth + "\t")

    # Then write all the todos for the project; and because each todo belongs to this specific project
    #   I don't have to add an area tag
    todos = get_all_todos_for_project_in_order(resolved_project)
    write_todos(todos, stream, depth=depth + "\t")


def write_projects(
    projects: list[UUID] | list[dict[str, Any]],
    stream: TextIO = sys.stdout,
    depth: str = "",
) -> None:
    """
    Write a list of projects, one at a time, to stream

    The projects are all indented to the same depth.  Because this function might also be called by an interactive
    user, it accepts a list of UUIDs, and has defaults for all other parameters.  It does its work by calling write_project
    which can take a UUID or a dict, so no testing has to happen in the loop.
    """
    for project in projects:
        write_project(project, stream, depth)


def get_all_projects_for_area(area: UUID | dict[str, Any]) -> list[dict[str, Any]]:
    """
    Extract all the projects that belong to the given area and return them as a list

    The return value is a list of dictionaries.  The list is sorted by index, so the projects should
    appear in the same order they have under the given area.  Interactive users might want that to be a list of UUIDs.
    That's easy to get: list_to_uuids(get_all_projects_for_area(area)).  Because this function might
    also be called by an interactive user, it accepts an area UUID.

    If there are no projects under this area, this function returns an empty list.
    """
    resolved_area = _resolve_item_if_needed(area, type="area")
    if not resolved_area:
        assert isinstance(area, str)
        raise UUIDDoesNotResolveError(
            area, "area"
        )  # this can only happen for an interactive user, because only interactive users use UUIDs

    area_uuid = resolved_area["uuid"]
    return sorted(
        [
            project
            for project in resolved_area.get("items", [])
            if project.get("type") == "project" and project.get("area") == area_uuid
        ],
        key=itemgetter("index"),
    )


def get_all_projects_with_no_area() -> list[dict[str, Any]]:
    """
    Extract all the projects that are outside of any area and return them as a list

    The return value is a list of dictionaries.  The list is sorted by index, so the projects
    should appear in the same order they are listed in the interface.  Interactive users might want that to be a list of UUIDs.
    That's easy to get: list_to_uuids(get_all_projects_with_no_area()).

    If there are no projects without areas, this function returns the empty list.
    """
    return sorted(
        [
            project
            for project in things.tasks(type="project", area=False)
            if project.get("type") == "project" and not project.get("area")
        ],
        key=itemgetter("index"),
    )


def get_all_todos_for_area(area: UUID | dict[str, Any]) -> list[dict[str, Any]]:
    """
    Extract a list of todos at the top-level of the given area and return them as a list

    The return value is a list of dictionaries.  The list is sorted by index, so the todos
    should appear in the same order they appear in the user interface.  Interactive users might want the result
    as a list of UUIDs.  That's easy to get: list_to_uuids(get_all_todos_for_area(area)).

    If there are no top-level todos in the given area, this function returns the empty list.
    """
    resolved_area = _resolve_item_if_needed(area, type="area")
    if not resolved_area:
        assert isinstance(area, str)
        raise UUIDDoesNotResolveError(
            area, "area"
        )  # this can only happen for an interactive user, because only interactive users use UUIDs

    area_uuid = resolved_area["uuid"]
    return sorted(
        [
            todo
            for todo in resolved_area.get("items", [])
            if todo.get("type") == "to-do" and todo.get("area") == area_uuid
        ],
        key=itemgetter("index"),
    )


def get_all_todos_with_no_area() -> list[dict[str, Any]]:
    """
    Extract a list of all todos not inside any area and return them as a list

    The return value is a list of dictionaries.  They are ordered by index, so they should apear
    in the same order they appear in the user interface.  An interactive user might want the result
    as a list of UUIDs.  That's easy to get: list_to_uuids(get_all_todos_with_no_area()).

    If there are no todos outside of any area, this function returns the empty list.
    """
    return sorted(
        [
            todo
            for todo in things.tasks(type="to-do", project=False, area=False, include_items=True)
            if not todo.get("area") and not todo.get("project") and not todo.get("heading")
        ],
        key=itemgetter("index"),
    )


def write_all_items_in_area(area: UUID | dict[str, Any], stream: TextIO = sys.stdout, depth: str = "") -> None:
    """
    Write everything inside an area, both projects and todos, to stream

    Does not write the area name (or I guess the area might have notes and maybe tags, too).  The reason is, though
    this is easily expressible in TaskPaper, it ruins the import into OmniFocus.  When areas are part of the data,
    no actual projects are created, regardless of how you express them.

    Because this can be called by an interactive user, it accepts a UUID for the area and all the other
    parameters have defaults.

    The items written are all indented to depth.
    """
    # Note: I considered having area is None mean all items in _no_ area, but I just don't like the idea;
    # and that case only happens once anyway, in _write_everything_except_areas
    resolved_area = _resolve_item_if_needed(area, type="area")
    if not resolved_area:
        assert isinstance(area, str)
        raise UUIDDoesNotResolveError(
            area, "area"
        )  # this can only happen for an interactive user, because only interactive users use UUIDs
    write_projects(get_all_projects_for_area(resolved_area), stream, depth=depth)
    write_todos(
        get_all_todos_for_area(resolved_area),
        stream,
        depth=depth,
        area_name=resolved_area["title"],
    )


def _write_everything_except_areas(stream: TextIO) -> None:
    """
    Export an entire Things database to TaskPaper format onto stream

    This is a private function.  Interactive users should use export instead.

    Starts at zero depth, which happens to be the defaults for all the write functions used here.
    """
    for area in sorted(things.areas(include_items=True), key=itemgetter("title")):
        # TODO: areas don't have indexes?  So write them out ordered by name.  Is this right?  Investigate.
        write_all_items_in_area(area, stream)
    write_projects(get_all_projects_with_no_area(), stream)
    write_todos(get_all_todos_with_no_area(), stream)


def export(
    # TODO: below, can I replace List and Optional with their modern equivalents?
    # TODO: should I add a parameter to send the output to a file?
    write_to_clipboard: Annotated[Optional[bool], typer.Option()] = False,
    database: Annotated[Optional[Path], typer.Option()] = None,
    uuids: Annotated[Optional[List[UUID]], typer.Argument()] = None,
) -> None:
    """
    This the CLI entry-point for converting from a Things 3 database to an OmniFocus ready TaskPaper file

    By default, the file is written to stdout, but if --write-to-clipboard, then instead of going to stdout,
    it is copied to the system clipboard.  If you have a Things database somewhere _other_ than the default
    location, you can optionally provide a path to it.  I don't know how that ever happens in practice, but
    we account for it.  With nothing else, export converts and outputs the entire database.  However, you
    can optionally supply a list of one or more UUIDs and then instead of outputting the world, it outputs
    just the items, one after the other, corresponding to those UUIDs.  The type of each UUID is deduced.
    """
    if database and database.is_file():
        things.pop_database(filepath=str(database.resolve()))

    stream = io.StringIO()

    if not uuids:
        _write_everything_except_areas(stream)
    else:
        for uuid in uuids:
            # Note: if there's more than one uuid given, we don't let a bad one stop us.  We go on to all the following uuids.
            try:
                item = uuid_to_item(uuid)
            except UUIDDoesNotResolveError:
                print(
                    f"Warning: UUID '{uuid}' was not found in the database",
                    file=sys.stderr,
                )
                continue
            type = item.get("type")
            match type:
                case "to-do":
                    todo = item
                    write_todo(todo, stream=stream)
                case "project":
                    project = item
                    write_project(project, stream=stream)
                case "area":
                    area = item
                    write_all_items_in_area(area, stream=stream)
                case _:
                    print(
                        f"Warning: the item with UUID '{uuid}' was not a 'to-do', 'project', or 'area'",
                        file=sys.stderr,
                    )
                    continue

    if write_to_clipboard:
        pyperclip.copy(stream.getvalue())  # ...so I can use OmniFocus automation to "Import TaskPaper from Clipboard"
    else:
        sys.stdout.write(stream.getvalue())  # ...so I can debug


if __name__ == "__main__":
    typer.run(export)
