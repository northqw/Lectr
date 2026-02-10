.PHONY: setup
setup:
	npm install

.PHONY: clean
clean:
	rm -rf dist release

.PHONY: build
build:
	npm run build

.PHONY: preview
preview:
	npm run preview

.PHONY: dev
dev:
	npm run dev

.PHONY: serve-dist
serve-dist:
	npm run serve-dist

.PHONY: build-serve
build-serve:
	npm run build-and-serve

.PHONY: app-dev
app-dev:
	npm run app:dev

.PHONY: app
app:
	npm run app

.PHONY: app-dist
app-dist:
	npm run app:dist
