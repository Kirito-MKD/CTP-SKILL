# CTP-SKILL

![СТП](https://img.shields.io/badge/%D0%A1%D0%A2%D0%9F-01--2024-blue)
![БГУИР](https://img.shields.io/badge/%D0%91%D0%93%D0%A3%D0%98%D0%A0-%D0%A4%D0%98%D0%91-darkgreen)
![проверено](https://img.shields.io/badge/%D0%BF%D1%80%D0%BE%D0%B2%D0%B5%D1%80%D0%B5%D0%BD%D0%BE-%D0%BA%D0%B0%D1%84%D0%B5%D0%B4%D1%80%D0%B0_%D0%98%D0%91_%C2%B7_%D0%9E%D0%9F%D0%9B%D0%A1%2C_%D0%A2%D0%9E%D0%98-brightgreen)

Оформление отчётов и курсовых БГУИР по СТП 01-2024 — навык для Claude и промт для других моделей. Проверено на кафедре информационной безопасности (ОПЛС, ТОИ).

| Файл | Для чего |
|---|---|
| `stp-bguir-oformlenie/` | навык для Claude |
| `PROMPT.md` | промт для Gemini, ChatGPT и др. |

## Claude

- **Claude Code** — скопировать папку `stp-bguir-oformlenie` в `~/.claude/skills/`.
- **claude.ai** — заархивировать папку в zip и загрузить в Settings → Capabilities → Skills.

## Другие модели

Вставить `PROMPT.md` первым сообщением (или в системные инструкции), затем — свой текст.

Модель сама спросит ФИО студента и преподавателя, кафедру и дисциплину.
