# SPDX-License-Identifier: MIT
# Copyright (c) 2026 JAEHYUK CHO
import asyncio
import json
import logging
from typing import Optional
from fastapi import WebSocket

logger = logging.getLogger(__name__)


class WSHub:
    """인메모리 WebSocket 브로드캐스터. 단일 프로세스 전용."""

    def __init__(self):
        # workspace_id → set of WebSocket
        self._rooms: dict[str, set[WebSocket]] = {}
        self._lock = asyncio.Lock()

    async def connect(self, workspace_id: str, ws: WebSocket):
        async with self._lock:
            if workspace_id not in self._rooms:
                self._rooms[workspace_id] = set()
            self._rooms[workspace_id].add(ws)
        logger.debug(f"WS connected: workspace={workspace_id}, total={len(self._rooms[workspace_id])}")

    async def disconnect(self, workspace_id: str, ws: WebSocket):
        async with self._lock:
            room = self._rooms.get(workspace_id)
            if room:
                room.discard(ws)
                if not room:
                    del self._rooms[workspace_id]
        logger.debug(f"WS disconnected: workspace={workspace_id}")

    async def broadcast(self, workspace_id: str, message: dict, exclude: Optional[WebSocket] = None):
        text = json.dumps(message, ensure_ascii=False)
        room = self._rooms.get(workspace_id, set())
        dead = set()
        for ws in list(room):
            if ws is exclude:
                continue
            try:
                await ws.send_text(text)
            except Exception:
                dead.add(ws)
        # cleanup dead connections
        if dead:
            async with self._lock:
                room = self._rooms.get(workspace_id, set())
                room -= dead
                if not room and workspace_id in self._rooms:
                    del self._rooms[workspace_id]

    def connection_count(self, workspace_id: str) -> int:
        return len(self._rooms.get(workspace_id, set()))


# Singleton
hub = WSHub()
