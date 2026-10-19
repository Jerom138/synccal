"""Contrat commun a tous les backends d'agenda.

Le moteur de reconciliation ne connait que cette interface. C'est elle qui rend
le projet insensible a la voie d'acces : basculer un agenda de la lecture seule
(flux ICS) a la lecture/ecriture (COM) ne change pas une ligne du moteur.
"""

from __future__ import annotations

from typing import Protocol, runtime_checkable

from core.model import CanonicalEvent, OriginRef, Window


class BackendError(RuntimeError):
    """Echec d'acces a un agenda.

    Toujours fatal pour le run en cours : un agenda illisible ne doit jamais
    etre confondu avec un agenda vide, sous peine de supprimer toutes ses copies.
    """


@runtime_checkable
class CalendarBackend(Protocol):
    key: str          # identifiant de l'agenda, tel que declare dans config.yaml
    label: str        # libelle lisible, utilise dans les rapports
    writable: bool    # False -> le moteur calcule le plan mais n'ecrit rien

    def fetch(self, window: Window) -> list[CanonicalEvent]:
        """Tous les evenements de la fenetre, copies comprises."""
        ...

    def create(self, source: CanonicalEvent, origin: OriginRef) -> None:
        """Cree une copie de `source` portant le marqueur `origin`."""
        ...

    def update(self, native_id: str, source: CanonicalEvent, origin: OriginRef) -> None:
        """Realigne une copie existante sur sa source."""
        ...

    def delete(self, native_id: str) -> None:
        """Supprime une copie devenue orpheline ou surnumeraire."""
        ...


class ReadOnlyBackend:
    """Base des backends en lecture seule : toute ecriture est une erreur de programmation."""

    writable = False

    def create(self, source: CanonicalEvent, origin: OriginRef) -> None:
        raise BackendError(f"{self.key}: agenda en lecture seule, creation impossible")

    def update(self, native_id: str, source: CanonicalEvent, origin: OriginRef) -> None:
        raise BackendError(f"{self.key}: agenda en lecture seule, modification impossible")

    def delete(self, native_id: str) -> None:
        raise BackendError(f"{self.key}: agenda en lecture seule, suppression impossible")
