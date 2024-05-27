import click

@click.command()
@click.version_option()
@click.argument(
    "say",
    type=click.Path(exists=True, file_okay=True, dir_okay=False, allow_dash=False),
)
def cli(say):
    print(say)