import pytest

from inferencia import APARELHOS, sugerir_causas


def test_formato_saida():
    r = sugerir_causas("screen cracked after drop", "Celular", k=3)
    assert len(r["sugestoes"]) == 3
    probs = [s["probabilidade"] for s in r["sugestoes"]]
    assert probs == sorted(probs, reverse=True) and all(0 <= p <= 1 for p in probs)
    assert sum(probs) <= 1.0 + 1e-6
    assert {"confianca_alta", "acao_recomendada", "cobertura_vocabulario"} <= set(r)


def test_casos_claros():
    assert sugerir_causas("screen cracked after drop", "Celular")["sugestoes"][0]["categoria_falha"] == "Tela"
    r = sugerir_causas("won't turn on, no charging light", "Notebook")
    assert r["sugestoes"][0]["categoria_falha"] == "Energia / bateria" and r["confianca_alta"]


@pytest.mark.parametrize("aparelho", APARELHOS + ("Desconhecido",))
def test_todos_os_aparelhos(aparelho):
    assert sugerir_causas("battery drains fast", aparelho)["sugestoes"]


@pytest.mark.parametrize("texto", ["", "   ", "1234 !!!", None, 42])
def test_texto_invalido(texto):
    with pytest.raises(ValueError):
        sugerir_causas(texto, "Celular")


def test_aparelho_invalido():
    with pytest.raises(ValueError):
        sugerir_causas("screen broken", "Xyz")


def test_k_invalido():
    with pytest.raises(ValueError):
        sugerir_causas("screen broken", "Celular", k=0)


def test_texto_sem_sentido_tem_confianca_baixa():
    r = sugerir_causas("asdf qwer zxcv", "Tablet")
    assert r["confianca_alta"] is False and r["cobertura_vocabulario"] == 0


def test_determinismo():
    a = sugerir_causas("phone is slow and freezes", "Celular")
    assert a == sugerir_causas("phone is slow and freezes", "Celular")
