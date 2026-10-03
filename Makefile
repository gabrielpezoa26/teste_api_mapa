COMPOSE = docker compose
SERVICE = app
GAP_M  ?= 30

.PHONY: all build run refresh gap shell down clean fclean re start-docker prune check-data

all: run

build:
	$(COMPOSE) build

run: build
	$(COMPOSE) run --rm $(SERVICE)

refresh: build
	$(COMPOSE) run --rm -e REFRESH=1 $(SERVICE)

gap: build
	$(COMPOSE) run --rm -e GAP_M=$(GAP_M) $(SERVICE)

shell: build
	$(COMPOSE) run --rm $(SERVICE) bash

down:
	$(COMPOSE) down --remove-orphans

clean: down
	$(COMPOSE) down --rmi all

fclean: clean
	rm -rf data

re: fclean all

start-docker:
	sudo systemctl start docker

prune:
	docker system prune -af --volumes

check-data:
	ls -la data
