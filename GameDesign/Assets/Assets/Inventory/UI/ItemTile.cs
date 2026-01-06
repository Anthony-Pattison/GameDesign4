using UnityEngine;
using UnityEngine.UI;
using TMPro;
using UnityEngine.EventSystems;

public class ItemTile : MonoBehaviour
{
    [field:SerializeField]
    public InventoryItem item {get; private set;}

    [SerializeField]
    TextMeshProUGUI _label;

    [SerializeField]
    TextMeshProUGUI _price;

    [SerializeField]
    Image _image;

    public float highlight {get; private set;}

    public void Display(InventoryItem item) {
        this.item = item;
        UpdateDisplay();
    }

    void OnValidate() {
        UpdateDisplay();
    }

    void UpdateDisplay() {
        if (item == null || item.displayName == _label.text)
            return;

        _image.sprite = item.sprite;
        _label.text = item.displayName;
        _price.text = $"${item.cost}";        
    }

    public void SetHighlight(float intensity) {
        highlight = intensity;
        float scale = 1f + intensity * 0.4f;
        _image.transform.localScale = new Vector3(scale, scale, scale);

        _label.color = new Color(1, 1, 1, 1 + (intensity - 1) * 0.2f);
    }
}
