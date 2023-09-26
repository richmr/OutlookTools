# SPDX-FileCopyrightText: 2023-present richmr <richmr@users.noreply.github.com>
#
# SPDX-License-Identifier: MIT
import click
import typer


from outlooktools.__about__ import __version__

app = typer.Typer(name="OutlookTools")

"""
@click.group(context_settings={"help_option_names": ["-h", "--help"]}, invoke_without_command=True)
@click.version_option(version=__version__, prog_name="OutlookTools")
def outlooktools():
    click.echo("Hello world!")
"""