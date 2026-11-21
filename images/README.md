# Иллюстрации к лекции «Шардирование и репликация»

Манифест файлов в `images/`. Все файлы — схемы/диаграммы (без стоковых фото).

| Файл | Что изображено | Источник (прямой URL) | Лицензия / атрибуция |
|---|---|---|---|
| `cap-theorem-triangle.svg` | CAP-теорема: треугольник Consistency / Availability / Partition tolerance | https://upload.wikimedia.org/wikipedia/commons/b/be/CAP_Theorem.svg (страница: https://commons.wikimedia.org/wiki/File:CAP_Theorem.svg) | CC BY-SA 4.0, автор Mooond. В оригинале опечатка «Availabilty». |
| `consistent-hashing-ring.png` | Кольцо consistent hashing: объекты → серверы/бакеты по кругу, следующий сервер по часовой стрелке | https://upload.wikimedia.org/wikipedia/commons/7/71/Consistent_Hashing_Sample_Illustration.png (страница: https://commons.wikimedia.org/wiki/File:Consistent_Hashing_Sample_Illustration.png) | CC BY-SA 4.0, автор WikiLinuz |
| `consistent-hashing-vnodes-dynamo.svg` | Consistent hashing с репликами (Dynamo): хеш-кольцо, диапазоны, реплики, vnodes | https://upload.wikimedia.org/wikipedia/commons/6/63/Dynamo_consistent_hashing.svg (страница: https://commons.wikimedia.org/wiki/File:Dynamo_consistent_hashing.svg) | CC BY-SA 3.0 (de), авторы Flash1984 / Marlus Gancher. Подписи на немецком. |
| `replication-master-slave.svg` | Репликация хранилища: Primary → Target, поток данных + Ack (синхронная/асинхронная) | https://upload.wikimedia.org/wikipedia/commons/a/ad/Storage_replication-en.svg (страница: https://commons.wikimedia.org/wiki/File:Storage_replication-en.svg) | CC BY-SA 4.0, автор Michel Bakni |
| `kafka-topic-partition-log.png` | Топик Kafka = набор партиций; партиция — упорядоченный лог с offset (пример шардирования) | https://kafka.apache.org/38/images/log_anatomy.png (страница: https://kafka.apache.org/38/documentation.html) | Apache License 2.0 (Apache Kafka) |
| `kafka-apis.png` | Архитектура Kafka: producers / brokers (партиции, реплики) / consumers | https://kafka.apache.org/38/images/kafka-apis.png | Apache License 2.0 (Apache Kafka) |
| `citus-architecture.png` | Citus: coordinator + worker-ноды, распределение шардов | https://raw.githubusercontent.com/citusdata/citus/main/images/citus-architecture.png (репозиторий: https://github.com/citusdata/citus) | Из официального репозитория Citus (Citus Data / Microsoft), атрибуция Citus Data |
| `citus-scale-out.png` | Citus: горизонтальное масштабирование с одной ноды на кластер | https://raw.githubusercontent.com/citusdata/citus/main/images/citus-scale-out.png | То же — Citus Data / Microsoft |
| `greenplum-mirrors.png` | Greenplum MPP: master, segment-хосты, primary и mirror-сегменты (high availability) | https://blogs.vmware.com/tanzu/wp-content/uploads/sites/154/2018/01/greenplum-mirrors.png (страница: https://blogs.vmware.com/tanzu/self-healing-greenplum-doctor-always/) | © VMware / Broadcom, блог VMware Tanzu. Не свободная лицензия — использовать в лекции с атрибуцией «VMware Tanzu blog». |
| `patroni-etcd-ha.svg` | Рекомендуемая HA-архитектура: PostgreSQL (Primary + Replica), Patroni, etcd, HAProxy, watchdog, pgBackRest | https://docs.percona.com/postgresql/15/_images/diagrams/ha-recommended.svg (страница: https://docs.percona.com/postgresql/15/solutions/ha-architecture.html) | CC BY 4.0 (документация Percona) |
| `patroni-etcd-ha-basic.svg` | Базовая HA-схема Patroni: PostgreSQL Primary/Replica + Patroni + etcd (3 ноды) | https://docs.percona.com/postgresql/15/_images/diagrams/HA-basic.svg | CC BY 4.0 (документация Percona) |
| `raft-log-replication.svg` | Raft: 1) выбор лидера (AppendEntries/heartbeat, RequestVote); 2) репликация журнала (committed/missing entries) | Самостоятельно нарисованная схема (исходника-URL нет) | CC0, на основе Ongaro & Ousterhout, «In Search of an Understandable Consensus Algorithm» (USENIX ATC 2014) |

## Замечания

- **Все целевые изображения скачаны успешно** — файлов, требующих ручного скачивания, нет.
- Raft: на Wikimedia Commons свободной диаграммы Raft (выборы лидера / репликация журнала) нет (есть только маскот «Raft Consensus Algorithm Mascot»), а рисунки из оригинальной статьи защищены авторским правом, поэтому `raft-log-replication.svg` нарисован вручную (CC0).
- Для печати/слайдов SVG-файлы (`*.svg`) желательно конвертировать в PNG/PDF нужного размера (например, `qlmanage -t` или `rsvg-convert`).
- Порядок проверки лицензий: Wikimedia Commons — CC BY-SA; Apache Kafka — Apache-2.0; Percona docs — CC BY 4.0; Citus и Greenplum — атрибуция без открытой лицензии на изображения.
