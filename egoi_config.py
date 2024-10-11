API_KEY = 'insert_api_key'
LIST_ID = '21'
BODY_EMAIL = """  
        <div class="comment_form">
        <h2>Dê-nos a sua opinião</h2>
        <p>
        <label for="star_rating">Como classifica este produto:</label>
        <select id="star_rating" name="star_rating">
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
        </select> <span class="st">(1-Mau 5-Óptimo)</span>
        </p>
        <p>
        <label for="resumo">Resuma o produto em duas ou três palavras</label>
        xtextarea name="resumo" style="height: 30px;">x/textarea>
        </p>
        <p>
        <label for="descricao_pros">O que mais gostou:</label>
        xtextarea name="descricao_pros" rows="4">x/textarea>
        </p>
        <p>
        <label for="descricao_contras">O que menos gostou:</label>
        xtextarea name="descricao_contras" rows="4">x/textarea>
        </p>
        <p>
        <label for="texto">Comentário final sobre o produto:</label>
        xtextarea name="texto" rows="4">x/textarea>
        </p>
        <p id="rec_prod">
        Recomendaria este produto a um amigo? <br>
        <div class="input-prepend">
            <span class="add-on"><input type="radio" value="1" name="recomenda" id="rec_sim" checked="true"></span>
            <label for="rec_sim" class="mini">Sim</label>
            <span class="add-on"><input type="radio" value="0" name="recomenda" id="rec_nao"></span>
            <label for="rec_nao" class="mini last">Não</label>
        </div>
        </p>
        <input type="submit" value="enviar" class="btn login_input" style="float:left;">
        </div>
        <div class="clear"></div>
        <hr>
        """
CAMPAIGN_HASH = 'insert_CAMPAIGN_HASH'
