using UnityEngine;
using UnityEngine.UI;

public class Dragger : MonoBehaviour
{
    [SerializeField] Image _image;
    [SerializeField] Image _shadow;

    [SerializeField] float _dropDistance;

    const float SHADOW_OPACITY = 1f;

    public void Select(Sprite sprite) {
        _image.sprite = sprite;        
        _image.color = Color.white;
        _shadow.color = default;
        transform.localScale = Vector3.zero;
    }

    public void Highlight(float amount) {
        float scale = 1f - Mathf.Clamp01(amount * 3f);
        scale = 1f - scale * scale * scale;
        transform.localScale = new Vector3(scale, scale, scale);
        _shadow.color = new Color(0, 0, 0, scale * SHADOW_OPACITY);
    }

    public bool DropAndDisable(float time) {
        float alpha = Mathf.Clamp01(1f - time / 0.8f);
        _image.color = new Color(1, 1, 1, alpha);
        _shadow.color = new Color(0, 0, 0, alpha * Mathf.Clamp01(1f - 5f * time) * SHADOW_OPACITY);

        alpha *= alpha;        
        transform.localScale = new Vector3(alpha, alpha, alpha);
        transform.Translate(0, -Time.deltaTime * time * _dropDistance, 0);
        return alpha == 0f;
    }
}
